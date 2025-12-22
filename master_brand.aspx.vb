Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports System.Web.UI.WebControls

Partial Public Class master_brand
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

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

            CheckMenuPermissions()
        End If
    End Sub

    Private Sub CheckMenuPermissions()

        Dim currentRole As String = Session("UserRole")


        If currentRole = "Buyer" Then
            grpmenuPO.Visible = False
            grpmenuMaster.Visible = False
            menuCreateOTBSwitching.Visible = False
            menuDraftOTBPlan.Visible = False
        Else
            grpmenuPO.Visible = True
            grpmenuMaster.Visible = True
            menuCreateOTBSwitching.Visible = True
            menuDraftOTBPlan.Visible = True
        End If


    End Sub

    ' Export to Excel
    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=MasterBrand_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".xls")
        Response.Charset = ""
        Response.ContentType = "application/vnd.ms-excel"

        Using sw As New StringWriter()
            Dim hw As New System.Web.UI.HtmlTextWriter(sw)

            ' Get data for export
            Dim dt As DataTable = GetBrandData()

            ' Create a temporary GridView for export
            Dim gvExport As New GridView()
            gvExport.DataSource = dt
            gvExport.DataBind()

            ' Style the header
            gvExport.HeaderRow.BackColor = System.Drawing.Color.LightGray
            For Each cell As TableCell In gvExport.HeaderRow.Cells
                cell.BackColor = System.Drawing.Color.FromArgb(11, 86, 164)
                cell.ForeColor = System.Drawing.Color.White
                cell.Font.Bold = True
            Next

            ' Render to HTML
            gvExport.RenderControl(hw)

            Response.Output.Write(sw.ToString())
            Response.Flush()
            Response.End()
        End Using
    End Sub

    ' Get Brand Data for Export
    Private Function GetBrandData() As DataTable
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT [Brand Code], [Brand Name], " &
                        "CASE WHEN ISNULL([isActive], 0) = 1 THEN 'Active' ELSE 'Inactive' END AS [Status] " &
                        "FROM [MS_Brand] ORDER BY [Brand Code]"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                conn.Open()
                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)
                Return dt
            End Using
        End Using
    End Function

    ' Override VerifyRenderingInServerForm for Export
    Public Overrides Sub VerifyRenderingInServerForm(control As Control)
        ' Required for export to Excel
    End Sub

End Class