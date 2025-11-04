Imports System
Imports System.ComponentModel.DataAnnotations
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.Script.Serialization
Imports System.Web.SessionState
Imports ExcelDataReader
Imports Newtonsoft.Json

Public Class POMatchingHandler
    Implements System.Web.IHttpHandler, IRequiresSessionState

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim action As String = If(context.Request("action"), "").ToLower().Trim()
        If action = "getpo" Then
            Try
                ' 1. รับค่า Parameters (ตัวอย่าง)
                Dim top As Integer = 1000
                Dim skip As Integer = 0

                ' 2. เรียก Helper ด้วย Workaround (Task.Run)
                Dim poList As List(Of SapPOResultItem) = Task.Run(Async Function()
                                                                      Return Await SapApiHelper.GetPOsAsync(Date.Today, top, skip)
                                                                  End Function).Result

                ' 3. ตรวจสอบผลลัพธ์
                If poList Is Nothing Then
                    Throw New Exception("Failed to get PO data from SAP.")
                End If

                ' 4. ส่งกลับเป็น JSON ให้ JavaScript (แนะนำวิธีนี้)
                context.Response.ContentType = "application/json"
                Dim successResponse = New With {
            .success = True,
            .count = poList.Count,
            .data = poList ' ส่ง List ทั้งหมดไปให้ JavaScript
        }
                context.Response.Write(JsonConvert.SerializeObject(successResponse))

            Catch ex As Exception
                context.Response.ContentType = "application/json"
                context.Response.StatusCode = 500
                Dim errorResponse As New With {
            .success = False,
            .message = ex.Message
        }
                context.Response.Write(JsonConvert.SerializeObject(errorResponse))
            End Try
        End If

    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class