Public Class Admin_matchPO
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Session("user") Is Nothing Then
            Response.Redirect("default.aspx")
        End If

        If Not IsPostBack Then
            ' 1. ตรวจสอบสิทธิ์
            Dim currentUser As String = Session("user").ToString()
            Dim currentPage As String = System.IO.Path.GetFileName(Request.Url.AbsolutePath)

            Dim rights As PermissionHelper.UserRights = PermissionHelper.GetPermission(currentUser, currentPage)

            ' 2. ถ้าไม่มีสิทธิ์ดู ให้ดีดออก
            'If Not rights.CanView Then
            '    Response.Write("<script>alert('คุณไม่มีสิทธิ์เข้าถึงหน้านี้'); window.location='dashboard.aspx';</script>")
            '    Response.End()
            'End If

        End If
    End Sub

End Class