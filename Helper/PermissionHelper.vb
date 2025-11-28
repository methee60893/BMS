Imports System.Data.SqlClient
Imports System.Web

Public Class PermissionHelper
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    ' Structure เก็บสิทธิ์ของผู้ใช้
    Public Structure UserRights
        Public CanView As Boolean
        Public CanEdit As Boolean
        Public CanDelete As Boolean
        Public CanApprove As Boolean
    End Structure

    ''' <summary>
    ''' ตรวจสอบสิทธิ์ของ User กับหน้าปัจจุบัน
    ''' </summary>
    Public Shared Function GetPermission(username As String, pageUrl As String) As UserRights
        Dim rights As New UserRights With {.CanView = False, .CanEdit = False, .CanDelete = False, .CanApprove = False}

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                ' Query นี้จะ Check ว่า User มี Role อะไร และ Role นั้นมีสิทธิ์ใน Menu นี้หรือไม่
                Dim sql As String = "
                    SELECT 
                        MAX(CAST(p.CanView AS INT)) AS CanView,
                        MAX(CAST(p.CanEdit AS INT)) AS CanEdit,
                        MAX(CAST(p.CanDelete AS INT)) AS CanDelete,
                        MAX(CAST(p.CanApprove AS INT)) AS CanApprove
                    FROM MS_User u
                    JOIN Map_User_Role ur ON u.UserID = ur.UserID
                    JOIN Map_Role_Permission p ON ur.RoleID = p.RoleID
                    JOIN MS_Menu m ON p.MenuID = m.MenuID
                    WHERE u.Username = @Username 
                      AND m.PageUrl = @PageUrl
                      AND u.IsActive = 1
                "
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Username", username)
                    cmd.Parameters.AddWithValue("@PageUrl", pageUrl)

                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.Read() AndAlso Not IsDBNull(reader("CanView")) Then
                            rights.CanView = Convert.ToBoolean(reader("CanView"))
                            rights.CanEdit = Convert.ToBoolean(reader("CanEdit"))
                            rights.CanDelete = Convert.ToBoolean(reader("CanDelete"))
                            rights.CanApprove = Convert.ToBoolean(reader("CanApprove"))
                        End If
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Log error
        End Try

        Return rights
    End Function
End Class