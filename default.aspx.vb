Imports System.Data.SqlClient
Imports BMS.share_class


Public Class _default
    Inherits System.Web.UI.Page

    Public connectionStringLocal As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Try
            Dim clientIP As String = GetClientIP()
            Dim email As String = txtEmail.Text.Trim()
            Dim password As String = txtPassword.Text
            Dim username As String = ""
            Dim fullName As String = email
            Dim userEmail As String = email

            If authen_by_ad(email, password) Then

                username = email

                Try
                    ' ดึงข้อมูล Profile จาก AD (ต้องใช้ Function ที่มีใน share_class)
                    Dim adProfile As share_class.retAD = share_class.GetLDAPUsersProfile(email)

                    ' เตรียมข้อมูล
                    Dim adFullName As String = (adProfile.ADname & " " & adProfile.ADSurname).Trim()
                    If Not String.IsNullOrWhiteSpace(adFullName) Then
                        fullName = adFullName
                    End If

                    If Not String.IsNullOrWhiteSpace(adProfile.ADemail) Then
                        userEmail = adProfile.ADemail
                    End If

                    ' บันทึกลงฐานข้อมูล
                    SyncUserWithDB(username, fullName, userEmail)

                Catch exSync As Exception
                    ' ถ้า Sync พลาด ให้ Log ไว้ แต่ยอมให้ Login ต่อไปได้ (Non-blocking)
                    System.Diagnostics.Trace.TraceWarning("Sync User Error: " & exSync.Message)
                End Try
                'If chkauthen Then

                Dim role = getRoleByUser(email)

                If role IsNot Nothing AndAlso role.Rows.Count > 0 Then

                    Session("UserRole") = role.Rows(0)("RoleName").ToString()
                Else
                    Session("UserRole") = "Buyer" ' ไม่มีบทบาท
                End If

                Session("Login") = True
                Session("user") = username
                Session("fullname") = fullName
                Response.Redirect("dashboard.aspx?u=" & Server.UrlEncode(email), False)
                Context.ApplicationInstance.CompleteRequest()
                Return
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
            System.Diagnostics.Trace.TraceWarning("Get role failed: " & ex.Message)
        End Try
        Return dt
    End Function



End Class
