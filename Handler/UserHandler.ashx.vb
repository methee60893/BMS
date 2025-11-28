Imports System
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.Script.Serialization
Imports System.Collections.Generic
Imports System.Configuration

Public Class UserHandler : Implements IHttpHandler
    Private connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString

    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.ContentType = "application/json"
        Dim action As String = context.Request("action")

        If action = "getUsers" Then
            GetUsers(context)
        ElseIf action = "saveUser" Then
            SaveUser(context)
        ElseIf action = "getRoles" Then
            GetRoles(context)
        End If
    End Sub

    Private Sub GetUsers(context As HttpContext)
        Dim search As String = context.Request("search")
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand("SP_Get_Users_List", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@SearchText", If(String.IsNullOrEmpty(search), DBNull.Value, search))
                conn.Open()
                dt.Load(cmd.ExecuteReader())
            End Using
        End Using

        ' Convert DataTable to List of Dictionary for JSON
        Dim rows As New List(Of Dictionary(Of String, Object))()
        For Each dr As DataRow In dt.Rows
            Dim row As New Dictionary(Of String, Object)()
            For Each col As DataColumn In dt.Columns
                row.Add(col.ColumnName, dr(col))
            Next
            rows.Add(row)
        Next
        context.Response.Write(New JavaScriptSerializer().Serialize(rows))
    End Sub

    Private Sub GetRoles(context As HttpContext)
        Dim dt As New DataTable()
        Using conn As New SqlConnection(connectionString)
            ' ดึง Role ทั้งหมด
            Using cmd As New SqlCommand("SELECT RoleID, RoleName FROM MS_Role", conn)
                conn.Open()
                dt.Load(cmd.ExecuteReader())
            End Using
        End Using
        ' Manual JSON construction for simplicity
        Dim rows As New List(Of Dictionary(Of String, Object))()
        For Each dr As DataRow In dt.Rows
            Dim row As New Dictionary(Of String, Object)()
            row.Add("id", dr("RoleID"))
            row.Add("text", dr("RoleName"))
            rows.Add(row)
        Next
        context.Response.Write(New JavaScriptSerializer().Serialize(rows))
    End Sub

    Private Sub SaveUser(context As HttpContext)
        Dim userId As Integer = Convert.ToInt32(context.Request.Form("userId"))
        Dim roleId As Integer = Convert.ToInt32(context.Request.Form("roleId"))
        Dim isActive As Boolean = Convert.ToBoolean(context.Request.Form("isActive"))

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand("SP_Admin_Save_User", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@UserID", userId)
                cmd.Parameters.AddWithValue("@RoleID", roleId)
                cmd.Parameters.AddWithValue("@IsActive", isActive)
                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
        context.Response.Write(New JavaScriptSerializer().Serialize(New With {.success = True, .message = "User updated successfully"}))
    End Sub

    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property
End Class