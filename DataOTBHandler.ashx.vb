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



    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class