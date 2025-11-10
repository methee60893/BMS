Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports System.Web.UI.WebControls

Partial Public Class master_vendor
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            LoadSegments()
            BindGridView()
        End If
    End Sub

    ' Load Segments for dropdown filter
    Private Sub LoadSegments()
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT DISTINCT [SegmentCode], [Segment] FROM [MS_Vendor] WHERE [SegmentCode] IS NOT NULL AND [SegmentCode] <> '' ORDER BY [Segment]"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                conn.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    Dim item As New ListItem(reader("Segment").ToString(), reader("SegmentCode").ToString())
                    ddlSearchSegment.Items.Add(item)
                End While
            End Using
        End Using
    End Sub

    ' Bind GridView with optional filters
    Private Sub BindGridView(Optional searchCode As String = "", Optional searchName As String = "", Optional segmentCode As String = "")
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT [VendorCode], [Vendor], [CCY], [PaymentTermCode], [PaymentTerm], [SegmentCode], [Segment], [Incoterm] FROM [MS_Vendor] WHERE 1=1"

        If Not String.IsNullOrEmpty(searchCode) Then
            query &= " AND [VendorCode] LIKE @Code"
        End If
        If Not String.IsNullOrEmpty(searchName) Then
            query &= " AND [Vendor] LIKE @Name"
        End If
        If Not String.IsNullOrEmpty(segmentCode) Then
            query &= " AND [SegmentCode] = @SegmentCode"
        End If

        query &= " ORDER BY [VendorCode]"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                If Not String.IsNullOrEmpty(searchCode) Then
                    cmd.Parameters.AddWithValue("@Code", "%" & searchCode & "%")
                End If
                If Not String.IsNullOrEmpty(searchName) Then
                    cmd.Parameters.AddWithValue("@Name", "%" & searchName & "%")
                End If
                If Not String.IsNullOrEmpty(segmentCode) Then
                    cmd.Parameters.AddWithValue("@SegmentCode", segmentCode)
                End If

                conn.Open()
                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)
                gvVendor.DataSource = dt
                gvVendor.DataBind()
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
        txtCreateCCY.Text = ""
        txtCreatePaymentTermCode.Text = ""
        txtCreatePaymentTerm.Text = ""
        txtCreateSegmentCode.Text = ""
        txtCreateSegment.Text = ""
        txtCreateIncoterm.Text = ""
    End Sub

    ' Create New Vendor
    Protected Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        Dim code As String = txtCreateCode.Text.Trim()
        Dim name As String = txtCreateName.Text.Trim()
        Dim ccy As String = txtCreateCCY.Text.Trim()
        Dim paymentTermCode As String = txtCreatePaymentTermCode.Text.Trim()
        Dim paymentTerm As String = txtCreatePaymentTerm.Text.Trim()
        Dim segmentCode As String = txtCreateSegmentCode.Text.Trim()
        Dim segment As String = txtCreateSegment.Text.Trim()
        Dim incoterm As String = txtCreateIncoterm.Text.Trim()

        ' Validation
        If String.IsNullOrEmpty(code) OrElse String.IsNullOrEmpty(name) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Vendor Code and Vendor Name are required!');", True)
            Return
        End If

        ' Check if Vendor Code already exists
        If CheckVendorCodeAndSectionExists(code, segmentCode) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Vendor Code With SegmentCode [" + segmentCode + "] already exists!');", True)
            Return
        End If

        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Using conn As New SqlConnection(connectionString)
            Dim query As String = "INSERT INTO MS_Vendor ([VendorCode], [Vendor], [CCY], [PaymentTermCode], [PaymentTerm], [SegmentCode], [Segment], [Incoterm]) " &
                                "VALUES (@code, @name, @ccy, @paymentTermCode, @paymentTerm, @segmentCode, @segment, @incoterm)"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", code)
                cmd.Parameters.AddWithValue("@name", name)
                cmd.Parameters.AddWithValue("@ccy", If(String.IsNullOrEmpty(ccy), DBNull.Value, DirectCast(ccy, Object)))
                cmd.Parameters.AddWithValue("@paymentTermCode", If(String.IsNullOrEmpty(paymentTermCode), DBNull.Value, DirectCast(paymentTermCode, Object)))
                cmd.Parameters.AddWithValue("@paymentTerm", If(String.IsNullOrEmpty(paymentTerm), DBNull.Value, DirectCast(paymentTerm, Object)))
                cmd.Parameters.AddWithValue("@segmentCode", If(String.IsNullOrEmpty(segmentCode), DBNull.Value, DirectCast(segmentCode, Object)))
                cmd.Parameters.AddWithValue("@segment", If(String.IsNullOrEmpty(segment), DBNull.Value, DirectCast(segment, Object)))
                cmd.Parameters.AddWithValue("@incoterm", If(String.IsNullOrEmpty(incoterm), DBNull.Value, DirectCast(incoterm, Object)))

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using

        ' Clear form and hide
        ClearCreateForm()
        createFormBox.Visible = False

        ' Reload segments and grid
        LoadSegments()
        BindGridView()

        ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Vendor created successfully!');", True)
    End Sub

    ' Check if Vendor Code exists
    Private Function CheckVendorCodeExists(vendorCode As String) As Boolean
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT COUNT(*) FROM MS_Vendor WHERE [VendorCode] = @code"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", vendorCode)
                conn.Open()
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    Private Function CheckVendorCodeAndSectionExists(vendorCode As String, segmentCode As String) As Boolean
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT COUNT(*) FROM MS_Vendor WHERE [VendorCode] = @code AND [SegmentCode] = @segmentCode"

        Using conn As New SqlConnection(connectionString)
            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@code", vendorCode)
                cmd.Parameters.AddWithValue("@segmentCode", segmentCode)
                conn.Open()
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return count > 0
            End Using
        End Using
    End Function

    ' View/Search Button
    Protected Sub btnView_Click(sender As Object, e As EventArgs) Handles btnView.Click
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim(), ddlSearchSegment.SelectedValue)
    End Sub

    ' Clear Filter Button
    Protected Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txtSearchCode.Text = ""
        txtSearchName.Text = ""
        ddlSearchSegment.SelectedIndex = 0
        BindGridView()
    End Sub

    ' GridView Row Editing
    Protected Sub gvVendor_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles gvVendor.RowEditing
        gvVendor.EditIndex = e.NewEditIndex
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim(), ddlSearchSegment.SelectedValue)
    End Sub

    ' GridView Row Updating
    Protected Sub gvVendor_RowUpdating(sender As Object, e As GridViewUpdateEventArgs) Handles gvVendor.RowUpdating
        Dim vendorCode As String = gvVendor.DataKeys(e.RowIndex).Value.ToString()

        ' Get updated values from textboxes in edit mode
        Dim vendor As String = DirectCast(gvVendor.Rows(e.RowIndex).Cells(1).Controls(0), TextBox).Text.Trim()
        Dim ccy As String = DirectCast(gvVendor.Rows(e.RowIndex).Cells(2).Controls(0), TextBox).Text.Trim()
        Dim paymentTermCode As String = DirectCast(gvVendor.Rows(e.RowIndex).Cells(3).Controls(0), TextBox).Text.Trim()
        Dim paymentTerm As String = DirectCast(gvVendor.Rows(e.RowIndex).Cells(4).Controls(0), TextBox).Text.Trim()
        Dim segmentCode As String = DirectCast(gvVendor.Rows(e.RowIndex).Cells(5).Controls(0), TextBox).Text.Trim()
        Dim segment As String = DirectCast(gvVendor.Rows(e.RowIndex).Cells(6).Controls(0), TextBox).Text.Trim()
        Dim incoterm As String = DirectCast(gvVendor.Rows(e.RowIndex).Cells(7).Controls(0), TextBox).Text.Trim()

        ' Validation
        If String.IsNullOrEmpty(vendor) Then
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Vendor Name is required!');", True)
            Return
        End If

        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Using conn As New SqlConnection(connectionString)
            Dim query As String = "UPDATE MS_Vendor SET [Vendor] = @vendor, [CCY] = @ccy, [PaymentTermCode] = @paymentTermCode, " &
                                "[PaymentTerm] = @paymentTerm, [SegmentCode] = @segmentCode, [Segment] = @segment, [Incoterm] = @incoterm " &
                                "WHERE [VendorCode] = @code"

            Using cmd As New SqlCommand(query, conn)
                cmd.Parameters.AddWithValue("@vendor", vendor)
                cmd.Parameters.AddWithValue("@ccy", If(String.IsNullOrEmpty(ccy), DBNull.Value, DirectCast(ccy, Object)))
                cmd.Parameters.AddWithValue("@paymentTermCode", If(String.IsNullOrEmpty(paymentTermCode), DBNull.Value, DirectCast(paymentTermCode, Object)))
                cmd.Parameters.AddWithValue("@paymentTerm", If(String.IsNullOrEmpty(paymentTerm), DBNull.Value, DirectCast(paymentTerm, Object)))
                cmd.Parameters.AddWithValue("@segmentCode", If(String.IsNullOrEmpty(segmentCode), DBNull.Value, DirectCast(segmentCode, Object)))
                cmd.Parameters.AddWithValue("@segment", If(String.IsNullOrEmpty(segment), DBNull.Value, DirectCast(segment, Object)))
                cmd.Parameters.AddWithValue("@incoterm", If(String.IsNullOrEmpty(incoterm), DBNull.Value, DirectCast(incoterm, Object)))
                cmd.Parameters.AddWithValue("@code", vendorCode)

                conn.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using

        gvVendor.EditIndex = -1
        LoadSegments()
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim(), ddlSearchSegment.SelectedValue)

        ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Vendor updated successfully!');", True)
    End Sub

    ' GridView Row Canceling Edit
    Protected Sub gvVendor_RowCancelingEdit(sender As Object, e As GridViewCancelEditEventArgs) Handles gvVendor.RowCancelingEdit
        gvVendor.EditIndex = -1
        BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim(), ddlSearchSegment.SelectedValue)
    End Sub

    ' GridView Row Deleting
    Protected Sub gvVendor_RowDeleting(sender As Object, e As GridViewDeleteEventArgs) Handles gvVendor.RowDeleting
        Dim vendorCode As String = gvVendor.DataKeys(e.RowIndex).Value.ToString()

        Try
            Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
            Using conn As New SqlConnection(connectionString)
                Dim query As String = "DELETE FROM MS_Vendor WHERE [VendorCode] = @code"

                Using cmd As New SqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@code", vendorCode)
                    conn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            LoadSegments()
            BindGridView(txtSearchCode.Text.Trim(), txtSearchName.Text.Trim(), ddlSearchSegment.SelectedValue)

            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Vendor deleted successfully!');", True)
        Catch ex As Exception
            ClientScript.RegisterStartupScript(Me.GetType(), "alert", "alert('Cannot delete this vendor. It may be referenced by other records.');", True)
        End Try
    End Sub

    ' Export to Excel
    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=MasterVendor_" & DateTime.Now.ToString("yyyyMMddHHmmss") & ".xls")
        Response.Charset = ""
        Response.ContentType = "application/vnd.ms-excel"

        Using sw As New StringWriter()
            Dim hw As New System.Web.UI.HtmlTextWriter(sw)

            ' Get data for export
            Dim dt As DataTable = GetVendorData()

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

    ' Get Vendor Data for Export
    Private Function GetVendorData() As DataTable
        Dim connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString").ConnectionString
        Dim query As String = "SELECT [VendorCode] AS 'Vendor Code', [Vendor] AS 'Vendor Name', [CCY], " &
                            "[PaymentTermCode] AS 'Payment Term Code', [PaymentTerm] AS 'Payment Term', " &
                            "[SegmentCode] AS 'Segment Code', [Segment], [Incoterm] " &
                            "FROM [MS_Vendor] ORDER BY [VendorCode]"

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