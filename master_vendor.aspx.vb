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