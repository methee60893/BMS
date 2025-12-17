Public Class actualPO
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
            If Not rights.CanView Then
                Response.Write("<script>alert('คุณไม่มีสิทธิ์เข้าถึงหน้านี้'); window.location='dashboard.aspx';</script>")
                Response.End()
            End If

            ' 3. ควบคุมปุ่มต่างๆ ตามสิทธิ์
            ' (สมมติหน้าจอมีปุ่ม btnSave, btnDelete, btnApprove)
            ' btnSave.Visible = rights.CanEdit
            ' btnDelete.Visible = rights.CanDelete
            ' btnApprove.Visible = rights.CanApprove
        End If
    End Sub

    Private Sub CheckMenuPermissions()
        ' ดึง Role ของ User ปัจจุบัน (ตัวอย่าง: ดึงจาก Session)
        Dim currentRole As String = Session("UserRole")

        ' หรือถ้าใช้ PermissionHelper ที่มีอยู่
        ' Dim hasApproveRight As Boolean = PermissionHelper.CheckPermission(CurrentUserID, "APPROVE_PO")

        ' --- ตัวอย่าง Logic การปิดเมนู ---

        ' กรณี: คนทั่วไป (User) -> ห้ามเห็นเมนูอนุมัติ
        If currentRole = "User" Then
            'menuApprovePO.Visible = False  ' ซ่อนเมนูนี้ทันที
        End If

        ' กรณี: ผู้บริหาร (Manager) -> เห็นได้ทุกเมนู
        If currentRole = "Manager" Then
            'menuApprovePO.Visible = True
            'menuCreatePO.Visible = True
        End If

        ' เทคนิค: ถ้า Logic ซับซ้อน ให้ Default เป็น False ไว้ก่อน แล้วเปิดเฉพาะที่มีสิทธิ์
        ' menuAdminPanel.Visible = False
        ' If PermissionHelper.IsAdmin(CurrentUserID) Then menuAdminPanel.Visible = True

    End Sub

End Class