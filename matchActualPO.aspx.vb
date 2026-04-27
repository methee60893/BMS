Public Class matchActualPO
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        PermissionHelper.EnsurePageAccess(Me)
    End Sub

End Class
