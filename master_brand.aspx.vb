Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports System.Web.UI.WebControls

Partial Public Class master_brand
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BindGridView()
        End If
    End Sub

    ' Bind GridView with optional filters
    Private Sub BindGridView(Optional searchCode As String = "", Optional searchName As String = "")
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT [Brand Code], [Brand Name] FROM [MS_Brand] WHERE 1=1"

        If Not String.IsNullOrEmpty(searchCode) Then
            query &= " AND [Brand Code] LIKE @Code"
        End If
        If Not String.IsNullOrEmpty(searchName) Then
            query &= " AND [Brand Name] LIKE @Name"
        End If

        query &= " ORDER BY [Brand Code]"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                If Not String.IsNullOrEmpty(searchCode) Then
                    cmd.Parameters.AddWithValue("@Code", "%" & searchCode & "%")
                End If
                If Not String.IsNullOrEmpty(searchName) Then
                    cmd.Parameters.AddWithValue("@Name", "%" & searchName & "%")
                End If

                conn.Open()
                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)
                gvBrand.DataSource = dt
                gvBrand.DataBind()
            End Using
        End Using
    End Sub

    ' Show Create Form
    Protected Sub btnShowCreate_Click(sender As Object, e As EventArgs) Handles btnShowCreate.Click
        ClearCreateForm()
        createFormBox.Visible = True
    End Sub

    ' Cancel Create
    Protected Sub btnCancelCreate_Click(sender As Object, e As EventArgs) Handles btnCancelCreate.Click
        ClearCreateForm()
        createFormBox.Visible = False
    End Sub

    ' Clear Create Form
    Private Sub ClearCreateForm()
        txtCreateCode.Text = ""
        txtCreateName.Text = ""
    End Sub

    ' Create New Brand
    Protected Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        Dim code As String = txtCreateCode.Text.Trim()
        Dim name As String = txtCreateName.Text.Trim()

        ' Validation
        If String.IsNullOrEmpty(code) OrElse String.IsNullOrEmpty(name) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Brand Code and Brand Name are required!');", True)
            Return
        End If

        ' Check if Brand Code already exists
        If CheckBrandCodeExists(code) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Brand Code already exists!');", True)
            Return
        End If

        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Using conn As New SqlConnection(connectionString)
            Dim query As String = "INSERT INTO MS_Brand ([Brand Code], [Brand Name]) VALUES (@code, @name)"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", code)
                cmd.Parameters.AddWithValue("@name", name)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' Clear form and hide
        ClearCreateForm()
        createFormBox.Visible = False

        ' Reload grid
        BindGridView()

        ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Brand created successfully!');", True)
    End Sub

    ' Check if Brand Code exists
    Private Function CheckBrandCodeExists(brandCode As String) As Boolean
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT COUNT(*) FROM MS_Brand WHERE [Brand Code] = @code"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", brandCode)
                conn.Open()
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    ' View/Search Button
    Protected Sub btnView_Click(sender As Object, e As EventArgs) Handles btnView.Click
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())
    End Sub

    ' Clear Filter Button
    Protected Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtSearchCode.Text = ""
        txtSearchName.Text = ""
        BindGridView()
    End Sub

    ' GridView Row Editing
    Protected Sub gvBrand_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles gvBrand.RowEditing
        gvBrand.EditIndex = e.NewEditIndex
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())
    End Sub

    ' GridView Row Updating
    Protected Sub gvBrand_RowUpdating(sender As Object, e As GridViewUpdateEventArgs) Handles gvBrand.RowUpdating
        Dim brandCode As String = gvBrand.DataKeys(e.RowIndex).Value.ToString()

        ' Get updated values from textboxes in edit mode
        Dim brandName As String = DirectCast(gvBrand.Rows(e.RowIndex).Cells(1).Controls(0), TextBox).Text.Trim()

        ' Validation
        If String.IsNullOrEmpty(brandName) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Brand Name is required!');", True)
            Return
        End If

        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Using conn As New SqlConnection(connectionString)
            Dim query As String = "UPDATE MS_Brand SET [Brand Name] = @name WHERE [Brand Code] = @code"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@name", brandName)
                cmd.Parameters.AddWithValue("@code", brandCode)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using

        gvBrand.EditIndex = -1
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())

        ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Brand updated successfully!');", True)
    End Sub

    ' GridView Row Canceling Edit
    Protected Sub gvBrand_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs) Handles gvBrand.RowCancelingEdit
        gvBrand.EditIndex = -1
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())
    End Sub

    ' GridView Row Deleting
    Protected Sub gvBrand_RowDeleting(sender As Object, e As GridViewDeleteEventArgs) Handles gvBrand.RowDeleting
        Dim brandCode As String = gvBrand.DataKeys(e.RowIndex).Value.ToString()

        Try
            Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
            Using conn As New SqlConnection(connectionString)
                Dim query As String = "DELETE FROM MS_Brand WHERE [Brand Code] = @code"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@code", brandCode)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())

            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Brand deleted successfully!');", True)
        Catch ex As Exception
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Cannot delete this brand. It may be referenced by other records.');", True)
        End Try
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
        Dim query As String = "SELECT [Brand Code], [Brand Name] FROM [MS_Brand] ORDER BY [Brand Code]"

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