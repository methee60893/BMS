Imports System.Data.SqlClient
Imports BMS.share_class


Public Class _default
    Inherits System.Web.UI.Page

    Public connectionString As String = ConfigurationManager.ConnectionStrings("LoginConnectionString")?.ConnectionString
    Public connectionStringLocal As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Try
            Dim clientIP As String = GetClientIP()
            Dim email As String = txtEmail.Text
            Dim password As String = txtPassword.Text
            Dim username As String = ""
            Dim fullName As String = ""
            Dim userEmail As String = ""

            If authen_by_ad(email, password) Then
                'Dim chkauthen = CheckADAccess(email)
                Try
                    ' ดึงข้อมูล Profile จาก AD (ต้องใช้ Function ที่มีใน share_class)
                    Dim adProfile As share_class.retAD = share_class.GetLDAPUsersProfile(email)

                    ' เตรียมข้อมูล
                    username = email
                    fullName = adProfile.ADname & " " & adProfile.ADSurname
                    userEmail = adProfile.ADemail

                    ' บันทึกลงฐานข้อมูล
                    SyncUserWithDB(username, fullName, userEmail)

                Catch exSync As Exception
                    ' ถ้า Sync พลาด ให้ Log ไว้ แต่ยอมให้ Login ต่อไปได้ (Non-blocking)
                    Console.WriteLine("Sync User Error: " & exSync.Message)
                End Try
                'If chkauthen Then

                Dim role = getRoleByUser(email)

                If role.Rows.Count > 0 Then

                    Session("UserRole") = role.Rows(0)("RoleName").ToString()
                Else
                    Session("UserRole") = "Buyer" ' ไม่มีบทบาท
                End If

                Session("Login") = True
                Session("user") = username
                Session("fullname") = fullName
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

    ' ฟังก์ชันใหม่สำหรับเรียก Stored Procedure
    Private Sub SyncUserWithDB(username As String, fullName As String, email As String)
        Using conn As New SqlConnection(connectionStringLocal)
            conn.Open()
            Using cmd As New SqlCommand("SP_Sync_User_From_AD", conn)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@Username", username)
                cmd.Parameters.AddWithValue("@FullName", If(String.IsNullOrEmpty(fullName), username, fullName))
                cmd.Parameters.AddWithValue("@Email", If(String.IsNullOrEmpty(email), "", email))
                cmd.ExecuteNonQuery()
            End Using
        End Using
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

    Public Function getRoleByUser(email As String) As DataTable
        Dim dt As New DataTable()
        Try
            Using conn As New SqlConnection(connectionStringLocal)
                Dim cmd As New SqlCommand("
                    SELECT 
                       [UserID]
                      ,[RoleID]
                      ,[RoleName]
                    FROM [BMS].[dbo].[View_UserRole]
                    WHERE Email = @Email
                ", conn)
                cmd.Parameters.AddWithValue("@Email", email)
                conn.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    dt.Load(reader)
                End Using
                conn.Close()
            End Using
        Catch ex As Exception
            dt = Nothing
        End Try
        Return dt
    End Function

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