Public Class manage_users
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        PermissionHelper.EnsurePageAccess(Me, adminOnly:=True)
    End Sub

End Class
