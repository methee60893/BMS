Imports System
Imports System.Web
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Text
Imports ExcelDataReader
Imports System.Globalization

Public Class DataOTBHandler
    Implements System.Web.IHttpHandler

    Public Shared connectionString93 As String = "Data Source=10.3.0.93;Initial Catalog=BMS;Persist Security Info=True;User ID=sa;Password=sql2014"

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        context.Response.Clear()
        context.Response.ContentType = "text/html"
        context.Response.ContentEncoding = Encoding.UTF8

        Try
            Dim dt As DataTable = Nothing


            If context.Request("action") = "obtlistbyfilter" Then

                Dim OTBtype As String = context.Request.Form("OTBtype").Trim()
                Dim OTByear As String = context.Request.Form("OTByear").Trim()
                Dim OTBmonth As String = context.Request.Form("OTBmonth").Trim()
                Dim OTBCompany As String = context.Request.Form("OTBCompany").Trim()
                Dim OTBCategory As String = context.Request.Form("OTBCategory").Trim()
                Dim OTBSegment As String = context.Request.Form("OTBSegment").Trim()
                Dim OTBBrand As String = context.Request.Form("OTBBrand").Trim()
                Dim OTBVendor As String = context.Request.Form("OTBVendor").Trim()
                dt = GetOTBDataWithFilter(OTBtype, OTByear, OTBmonth, OTBCompany, OTBCategory, OTBSegment, OTBBrand, OTBVendor)
                context.Response.Write(GenerateHtmlTable(dt))
            ElseIf context.Request("action") = "obtlistbyfilter" Then
                dt = GetOTBData()
                context.Response.Write(GenerateHtmlTable(dt))
            End If

        Catch ex As Exception
            context.Response.StatusCode = 500
            context.Response.Write($"<div class='alert alert-danger'>Error: {HttpUtility.HtmlEncode(ex.Message)}</div>")
        End Try

    End Sub

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

    Private Function GetOTBData() As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString93)
            conn.Open()
            Dim query As String = "SELECT * FROM [Template_Upload_Draft_OTB]" ' ปรับเปลี่ยนตามตารางจริง
            Using cmd As New SqlCommand(query, conn)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Private Function GetOTBDataWithFilter(ByVal OTBtype As String, ByVal OTByear As String, ByVal OTBmonth As String, ByVal OTBCompany As String, ByVal OTBCategory As String, ByVal OTBSegment As String, ByVal OTBBrand As String, ByVal OTBVendor As String) As DataTable
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString93)
            conn.Open()
            Dim query As String = "SELECT * FROM [Template_Upload_Draft_OTB]
                                    WHERE (@OTBtype = '' OR OTBType = @OTBtype)
                                      AND (@OTByear = '' OR OTBYear = @OTByear)
                                      AND (@OTBmonth = '' OR OTBMonth = @OTBmonth)
                                      AND (@OTBCompany = '' OR OTBCompany = @OTBCompany)
                                      AND (@OTBCategory = '' OR OTBCategory = @OTBCategory)
                                      AND (@OTBSegment = '' OR OTBSegment = @OTBSegment)
                                      AND (@OTBBrand = '' OR OTBBrand = @OTBBrand)
                                      AND (@OTBVendor = '' OR OTBVendor = @OTBVendor)
                                    " ' ปรับเปลี่ยนตามตารางจริง
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@OTBtype", OTBtype)
                cmd.Parameters.AddWithValue("@OTByear", OTByear)
                cmd.Parameters.AddWithValue("@OTBmonth", OTBmonth)
                cmd.Parameters.AddWithValue("@OTBCompany", OTBCompany)
                cmd.Parameters.AddWithValue("@OTBCategory", OTBCategory)
                cmd.Parameters.AddWithValue("@OTBSegment", OTBSegment)
                cmd.Parameters.AddWithValue("@OTBBrand", OTBBrand)
                cmd.Parameters.AddWithValue("@OTBVendor", OTBVendor)
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
            End Using
        End Using
        Return dt
    End Function


    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class