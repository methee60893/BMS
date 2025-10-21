Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports System.Web.UI.WebControls

Partial Public Class master_category
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BindGridView()
        End If
    End Sub

    ' Bind GridView with optional filters
    Private Sub BindGridView(Optional searchCode As String = "", Optional searchName As String = "")
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT [Cate], [Category] FROM [MS_Category] WHERE 1=1"

        If Not String.IsNullOrEmpty(searchCode) Then
            query &= " AND [Cate] LIKE @Code"
        End If
        If Not String.IsNullOrEmpty(searchName) Then
            query &= " AND [Category] LIKE @Name"
        End If

        query &= " ORDER BY [Cate]"

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
                gvCategory.DataSource = dt
                gvCategory.DataBind()
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

    ' Create New Category
    Protected Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        Dim code As String = txtCreateCode.Text.Trim()
        Dim name As String = txtCreateName.Text.Trim()

        ' Validation
        If String.IsNullOrEmpty(code) OrElse String.IsNullOrEmpty(name) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Category Code and Category Name are required!');", True)
            Return
        End If

        ' Check if Category Code already exists
        If CheckCategoryCodeExists(code) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Category Code already exists!');", True)
            Return
        End If

        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Using conn As New SqlConnection(connectionString)
            Dim query As String = "INSERT INTO MS_Category ([Cate], [Category]) VALUES (@code, @name)"

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

        ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Category created successfully!');", True)
    End Sub

    ' Check if Category Code exists
    Private Function CheckCategoryCodeExists(categoryCode As String) As Boolean
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT COUNT(*) FROM MS_Category WHERE [Cate] = @code"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", categoryCode)
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
    Protected Sub gvCategory_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles gvCategory.RowEditing
        gvCategory.EditIndex = e.NewEditIndex
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())
    End Sub

    ' GridView Row Updating
    Protected Sub gvCategory_RowUpdating(sender As Object, e As GridViewUpdateEventArgs) Handles gvCategory.RowUpdating
        Dim categoryCode As String = gvCategory.DataKeys(e.RowIndex).Value.ToString()

        ' Get updated values from textboxes in edit mode
        Dim categoryName As String = DirectCast(gvCategory.Rows(e.RowIndex).Cells(1).Controls(0), TextBox).Text.Trim()

        ' Validation
        If String.IsNullOrEmpty(categoryName) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Category Name is required!');", True)
            Return
        End If

        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Using conn As New SqlConnection(connectionString)
            Dim query As String = "UPDATE MS_Category SET [Category] = @name WHERE [Cate] = @code"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@name", categoryName)
                cmd.Parameters.AddWithValue("@code", categoryCode)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using

        gvCategory.EditIndex = -1
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())

        ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Category updated successfully!');", True)
    End Sub

    ' GridView Row Canceling Edit
    Protected Sub gvCategory_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs) Handles gvCategory.RowCancelingEdit
        gvCategory.EditIndex = -1
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())
    End Sub

    ' GridView Row Deleting
    Protected Sub gvCategory_RowDeleting(sender As Object, e As GridViewDeleteEventArgs) Handles gvCategory.RowDeleting
        Dim categoryCode As String = gvCategory.DataKeys(e.RowIndex).Value.ToString()

        Try
            Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
            Using conn As New SqlConnection(connectionString)
                Dim query As String = "DELETE FROM MS_Category WHERE [Cate] = @code"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@code", categoryCode)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim())

            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Category deleted successfully!');", True)
        Catch ex As Exception
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Cannot delete this category. It may be referenced by other records.');", True)
        End Try
    End Sub

    ' Export to Excel
    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=MasterCategory_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".xls")
        Response.Charset = ""
        Response.ContentType = "application/vnd.ms-excel"

        Using sw As New StringWriter()
            Dim hw As New System.Web.UI.HtmlTextWriter(sw)

            ' Get data for export
            Dim dt As DataTable = GetCategoryData()

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

    ' Get Category Data for Export
    Private Function GetCategoryData() As DataTable
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT [Cate] AS 'Category Code', [Category] AS 'Category Name' FROM [MS_Category] ORDER BY [Cate]"

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