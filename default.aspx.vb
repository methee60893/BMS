Imports System.Data.SqlClient
Imports BMS.share_class


Public Class _default
    Inherits System.Web.UI.Page

    Public connectionString As String = ConfigurationManager.ConnectionStrings("LoginConnectionString")?.ConnectionString

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Try
            Dim clientIP As String = GetClientIP()
            Dim email As String = txtEmail.Text
            Dim password As String = txtPassword.Text

            'adddblog()

            If authen_by_ad(email, password) Then
                'Dim chkauthen = CheckADAccess(email)
                'If chkauthen Then
                Session("Login") = True
                Session("user") = email
                Response.Redirect("dashboard.aspx?u=" & Server.UrlEncode(email))
                'Else
                '    ClientScript.RegisterStartupScript(Me.GetType(), "AuthorizationFailed", "alert(""You don't have authorization!"");", True)
                'End If
            ElseIf email = "admin" And password = "admin" Then
                'Session("Login") = True
                'Response.Redirect("camera.aspx?u=" & Server.UrlEncode(email))
            Else
                ClientScript.RegisterStartupScript(Me.GetType(), "LoginFailed", "alert(""Username หรือ Password ไม่ถูกต้อง!"");", True)
            End If
        Catch ex As Exception
            Dim safeMsg As String = ex.Message.Replace("""", "\""").Replace(vbCrLf, "\n")
            ClientScript.RegisterStartupScript(Me.GetType(), "Error", $"alert(""เกิดข้อผิดพลาด: {safeMsg}"");", True)
        End Try
    End Sub

    Private Sub adddblog()
        Try
            Try


                Using connection As New SqlConnection(connectionString)
                    connection.Open()

                    Dim sql As String = "INSERT INTO [dbo].[TB_Master_Log_Transection] ([ProjectName], [UserHostName], [UserHostAddress], [KeyWord], [CreateDateTime]) " &
                                        "VALUES (@ProjectName, @UserHostName, @UserHostAddress, @KeyWord, @CreateDateTime)"

                    Using command As New SqlCommand(sql, connection)
                        command.Parameters.AddWithValue("@ProjectName", "ONE SYSTEM")
                        command.Parameters.AddWithValue("@UserHostName", HttpContext.Current.Request.UserHostName)
                        command.Parameters.AddWithValue("@UserHostAddress", HttpContext.Current.Request.UserHostAddress)
                        command.Parameters.AddWithValue("@KeyWord", txtEmail.Text)
                        command.Parameters.AddWithValue("@CreateDateTime", DateTime.Now)

                        command.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try

        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Public Function CheckADAccess(username As String) As Boolean

        ' ===== ตรวจสอบสิทธิ์ในฐานข้อมูล =====
        Try
            Using conn As New SqlConnection(connectionString)
                Dim cmd As New SqlCommand("
                    SELECT CanAccessIndex 
                    FROM UserAccess 
                    WHERE Username = @Username
                ", conn)
                cmd.Parameters.AddWithValue("@Username", username)

                conn.Open()
                Dim result As Object = cmd.ExecuteScalar()
                conn.Close()

                If result IsNot Nothing AndAlso Convert.ToBoolean(result) = True Then
                    Return True
                End If
            End Using
        Catch ex As Exception
            Return False
        End Try

        Return False
    End Function

End Class