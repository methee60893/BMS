Imports System.Configuration
Imports System.Data.SqlClient
Imports System.IO
Imports System.Web.UI
Imports System.Collections.Generic

Public Class PermissionHelper
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Private Shared ReadOnly MenuPageByControlId As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
        {"menuDraftOTBPlan", "draftOTB.aspx"},
        {"menuApprovedOTBPlan", "approvedOTB.aspx"},
        {"menuCreateOTBSwitching", "createOTBswitching.aspx"},
        {"menuSwitchingTransaction", "transactionOTBSwitching.aspx"},
        {"menuCreateDraftPO", "createDraftPO.aspx"},
        {"menuDraftPO", "draftPO.aspx"},
        {"menuMatchActualPO", "matchActualPO.aspx"},
        {"menuActualPO", "actualPO.aspx"},
        {"menuOTBRemaining", "otbRemaining.aspx"},
        {"menuVendor", "master_vendor.aspx"},
        {"menuBrand", "master_brand.aspx"},
        {"menuCategory", "master_category.aspx"},
        {"menuAdminMatchPO", "admin_matchPO.aspx"},
        {"menuManageUsers", "manage_users.aspx"}
    }

    Private Shared ReadOnly AdminOnlyPages As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "admin_matchPO.aspx",
        "manage_users.aspx"
    }

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
                    cmd.Parameters.AddWithValue("@PageUrl", NormalizePageUrl(pageUrl))

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

    Public Shared Function GetPermission(username As String, pageUrl As String, roleName As Object) As UserRights
        Dim normalizedPage As String = NormalizePageUrl(pageUrl)

        If IsAdminOnlyPage(normalizedPage) Then
            If IsAdminRole(roleName) Then
                Return New UserRights With {.CanView = True, .CanEdit = True, .CanDelete = True, .CanApprove = True}
            End If

            Return New UserRights With {.CanView = False, .CanEdit = False, .CanDelete = False, .CanApprove = False}
        End If

        Return GetPermission(username, normalizedPage)
    End Function

    Public Shared Function IsAdminRole(roleName As Object) As Boolean
        If roleName Is Nothing Then
            Return False
        End If

        Return String.Equals(Convert.ToString(roleName).Trim(), "Admin", StringComparison.OrdinalIgnoreCase)
    End Function

    Public Shared Sub ApplyMenuPermissions(page As Page, username As String, roleName As Object)
        If page Is Nothing Then
            Return
        End If

        Dim viewablePages As HashSet(Of String) = GetViewableMenuPages(username, roleName)
        Dim visibleByControlId As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)

        For Each menuEntry In MenuPageByControlId
            Dim pageUrl As String = NormalizePageUrl(menuEntry.Value)
            Dim isVisible As Boolean = viewablePages.Contains(pageUrl)

            If IsAdminOnlyPage(pageUrl) Then
                isVisible = IsAdminRole(roleName)
            End If

            visibleByControlId(menuEntry.Key) = isVisible
            SetControlVisible(page, menuEntry.Key, isVisible)
        Next

        SetGroupVisible(page, visibleByControlId, "grpmenuOTBPlan", "menuDraftOTBPlan", "menuApprovedOTBPlan")
        SetGroupVisible(page, visibleByControlId, "grpmenuOTBSwitching", "menuCreateOTBSwitching", "menuSwitchingTransaction")
        SetGroupVisible(page, visibleByControlId, "grpmenuPO", "menuCreateDraftPO", "menuDraftPO", "menuMatchActualPO", "menuActualPO")
        SetGroupVisible(page, visibleByControlId, "grpmenuMaster", "menuVendor", "menuBrand", "menuCategory")
        SetGroupVisible(page, visibleByControlId, "grpmenuAdmin", "menuAdminMatchPO", "menuManageUsers")
    End Sub

    Public Shared Function EnsurePageAccess(page As Page, Optional adminOnly As Boolean = False) As UserRights
        If page Is Nothing Then
            Return New UserRights()
        End If

        If page.Session("user") Is Nothing Then
            page.Response.Redirect("default.aspx")
            page.Response.End()
        End If

        Dim currentUser As String = Convert.ToString(page.Session("user"))
        Dim currentRole As Object = page.Session("UserRole")
        Dim currentPage As String = NormalizePageUrl(page.Request.Url.AbsolutePath)

        If adminOnly AndAlso Not IsAdminRole(currentRole) Then
            DenyPage(page)
        End If

        Dim rights As UserRights = GetPermission(currentUser, currentPage, currentRole)
        If Not rights.CanView Then
            DenyPage(page)
        End If

        ApplyMenuPermissions(page, currentUser, currentRole)
        Return rights
    End Function

    Private Shared Function GetViewableMenuPages(username As String, roleName As Object) As HashSet(Of String)
        Dim pages As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        If String.IsNullOrWhiteSpace(username) Then
            Return pages
        End If

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Dim sql As String = "
                    SELECT DISTINCT m.PageUrl
                    FROM MS_User u
                    JOIN Map_User_Role ur ON u.UserID = ur.UserID
                    JOIN Map_Role_Permission p ON ur.RoleID = p.RoleID
                    JOIN MS_Menu m ON p.MenuID = m.MenuID
                    WHERE u.Username = @Username
                      AND u.IsActive = 1
                      AND ISNULL(m.IsActive, 1) = 1
                      AND p.CanView = 1
                "

                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@Username", username)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            If Not IsDBNull(reader("PageUrl")) Then
                                pages.Add(NormalizePageUrl(Convert.ToString(reader("PageUrl"))))
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            ' Fail closed: if permissions cannot be loaded, menus stay hidden except admin-only fallback.
        End Try

        If IsAdminRole(roleName) Then
            For Each adminPage In AdminOnlyPages
                pages.Add(adminPage)
            Next
        End If

        Return pages
    End Function

    Private Shared Function IsAdminOnlyPage(pageUrl As String) As Boolean
        Return AdminOnlyPages.Contains(NormalizePageUrl(pageUrl))
    End Function

    Private Shared Function NormalizePageUrl(pageUrl As String) As String
        If String.IsNullOrWhiteSpace(pageUrl) Then
            Return String.Empty
        End If

        Dim cleanUrl As String = pageUrl.Split("?"c)(0).Replace("\"c, "/"c)
        Return Path.GetFileName(cleanUrl)
    End Function

    Private Shared Sub DenyPage(page As Page)
        page.Response.Write("<script>alert('คุณไม่มีสิทธิ์เข้าถึงหน้านี้'); window.location='dashboard.aspx';</script>")
        page.Response.End()
    End Sub

    Private Shared Sub SetControlVisible(page As Page, controlId As String, visible As Boolean)
        Dim control As Control = FindControlRecursive(page, controlId)
        If control IsNot Nothing Then
            control.Visible = visible
        End If
    End Sub

    Private Shared Sub SetGroupVisible(page As Page, visibleByControlId As Dictionary(Of String, Boolean), groupId As String, ParamArray childIds As String())
        Dim groupVisible As Boolean = False

        For Each childId As String In childIds
            If visibleByControlId.ContainsKey(childId) AndAlso visibleByControlId(childId) Then
                groupVisible = True
                Exit For
            End If
        Next

        SetControlVisible(page, groupId, groupVisible)
    End Sub

    Private Shared Function FindControlRecursive(parent As Control, controlId As String) As Control
        If parent Is Nothing Then
            Return Nothing
        End If

        Dim found As Control = parent.FindControl(controlId)
        If found IsNot Nothing Then
            Return found
        End If

        For Each child As Control In parent.Controls
            found = FindControlRecursive(child, controlId)
            If found IsNot Nothing Then
                Return found
            End If
        Next

        Return Nothing
    End Function
End Class
