Imports System.Web
Imports System.Web.Services

Public Class DataPOHandler
    Implements System.Web.IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        Dim action As String = If(context.Request("action"), "").ToLower().Trim()

        If action = "exportdraftotb" Then
            Dim draftPONo As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftPONo")), "", context.Request.QueryString("draftPONo").Trim())
            Dim draftYear As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftYear")), "", context.Request.QueryString("draftYear").Trim())
            Dim draftMonth As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftMonth")), "", context.Request.QueryString("draftMonth").Trim())
            Dim draftCategory As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCategory")), "", context.Request.QueryString("draftCategory").Trim())
            Dim draftCompany As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCompany")), "", context.Request.QueryString("draftCompany").Trim())
            Dim draftBrand As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftBrand")), "", context.Request.QueryString("draftBrand").Trim())
            Dim draftVendor As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftVendor")), "", context.Request.QueryString("draftVendor").Trim())
            Dim draftAmountTHB As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftAmountTHB")), "", context.Request.QueryString("draftAmountTHB").Trim())
            Dim draftAmountCCY As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftAmountCCY")), "", context.Request.QueryString("draftAmountCCY").Trim())
            Dim draftCCY As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCCY")), "", context.Request.QueryString("draftCCY").Trim())
            Dim draftExhangeRate As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftExhangeRate")), "", context.Request.QueryString("draftExhangeRate").Trim())
            Dim draftRemark As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftRemark")), "", context.Request.QueryString("draftRemark").Trim())

            ' 1. Get Raw Data
            Dim dtExport As New DataTable("DraftPO")
            dtExport.Columns.Add("draftPODate", GetType(String))
            dtExport.Columns.Add("draftPONo", GetType(String))
            dtExport.Columns.Add("draftType", GetType(String))
            dtExport.Columns.Add("draftYear", GetType(String))
            dtExport.Columns.Add("draftMonth", GetType(String))
            dtExport.Columns.Add("draftCategory", GetType(String))
            dtExport.Columns.Add("CategoryName", GetType(String))
            dtExport.Columns.Add("draftCompany", GetType(String))
            dtExport.Columns.Add("CompanyName", GetType(String))
            dtExport.Columns.Add("draftSegment", GetType(String))
            dtExport.Columns.Add("SegmentName", GetType(String))
            dtExport.Columns.Add("draftBrand", GetType(String))
            dtExport.Columns.Add("BrandName", GetType(String))
            dtExport.Columns.Add("draftVendor", GetType(String))
            dtExport.Columns.Add("VendorName", GetType(String))
            dtExport.Columns.Add("AmountTHB", GetType(Decimal))
            dtExport.Columns.Add("AmountCCY", GetType(Decimal))
            dtExport.Columns.Add("CCY", GetType(String))
            dtExport.Columns.Add("ExRate", GetType(Decimal))
            dtExport.Columns.Add("ActualPORef", GetType(String))
            dtExport.Columns.Add("Status", GetType(String))
            dtExport.Columns.Add("StatusDate", GetType(String))
            dtExport.Columns.Add("Remark", GetType(String))
            dtExport.Columns.Add("ActionBy", GetType(String))

        Else
            context.Response.Clear()
            context.Response.ContentType = "text/html"
            context.Response.ContentEncoding = Encoding.UTF8
            Dim dt As DataTable = Nothing

            Dim draftPONo As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftPONo")), "", context.Request.QueryString("draftPONo").Trim())
            Dim draftYear As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftYear")), "", context.Request.QueryString("draftYear").Trim())
            Dim draftMonth As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftMonth")), "", context.Request.QueryString("draftMonth").Trim())
            Dim draftCategory As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCategory")), "", context.Request.QueryString("draftCategory").Trim())
            Dim draftCompany As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCompany")), "", context.Request.QueryString("draftCompany").Trim())
            Dim draftBrand As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftBrand")), "", context.Request.QueryString("draftBrand").Trim())
            Dim draftVendor As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftVendor")), "", context.Request.QueryString("draftVendor").Trim())
            Dim draftAmountTHB As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftAmountTHB")), "", context.Request.QueryString("draftAmountTHB").Trim())
            Dim draftAmountCCY As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftAmountCCY")), "", context.Request.QueryString("draftAmountCCY").Trim())
            Dim draftCCY As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftCCY")), "", context.Request.QueryString("draftCCY").Trim())
            Dim draftExhangeRate As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftExhangeRate")), "", context.Request.QueryString("draftExhangeRate").Trim())
            Dim draftRemark As String = If(String.IsNullOrWhiteSpace(context.Request.QueryString("draftRemark")), "", context.Request.QueryString("draftRemark").Trim())

            If context.Request("action") = "PoListFilter" Then

            ElseIf context.Request("action") = "PoListPreview" Then

            ElseIf context.Request("action") = "InsertPreview" Then

            ElseIf context.Request("action") = "PreviewEdit" Then

            ElseIf context.Request("action") = "EditDraftPO" Then

            End If
        End If
    End Sub

    Private Function GenerateHtmlApprovedTable(dt As DataTable) As String
        Dim sb As New StringBuilder()
        If dt.Rows.Count = 0 Then
            sb.Append("<tr><td colspan='24' class='text-center text-muted'>No approved OTB records found</td></tr>")
        Else
            For Each row As DataRow In dt.Rows
                sb.Append("<tr>")

                ' 1. Create Date (เช่น draftPODate หรือ CreateDT)
                Dim createDate As String = ""
                If row.Table.Columns.Contains("draftPODate") AndAlso row("draftPODate") IsNot DBNull.Value Then
                    createDate = Convert.ToDateTime(row("draftPODate")).ToString("dd/MM/yyyy HH:mm tt")
                ElseIf row.Table.Columns.Contains("CreateDT") AndAlso row("CreateDT") IsNot DBNull.Value Then
                    createDate = Convert.ToDateTime(row("CreateDT")).ToString("dd/MM/yyyy HH:mm tt")
                End If
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(createDate))

                ' 2. PO No (draftPONo)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftPONo")))

                ' 3. Status/Type (draftType)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftType")))

                ' 4. Year (draftYear)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftYear")))

                ' 5. Month (draftMonth - แปลงจากตัวเลขเป็นชื่อย่อ)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetMonthName(GetDbString(row, "draftMonth"))))

                ' 6. Cat Code (draftCategory)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftCategory")))

                ' 7. Cat Name (CategoryName)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "CategoryName")))

                ' 8. Comp Code (draftCompany)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftCompany")))

                ' 9. Seg Code (draftSegment)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftSegment")))

                ' 10. Seg Name (SegmentName)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "SegmentName")))

                ' 11. Brand Code (draftBrand)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftBrand")))

                ' 12. Brand Name (BrandName)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "BrandName")))

                ' 13. Vendor Code (draftVendor)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "draftVendor")))

                ' 14. Vendor Name (VendorName)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "VendorName")))

                ' 15. Amount THB (AmountTHB)
                sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "AmountTHB").ToString("N2"))

                ' 16. Amount CCY (AmountCCY)
                sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "AmountCCY").ToString("N2"))

                ' 17. Currency (CCY)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "CCY")))

                ' 18. Ex. Rate (ExRate)
                sb.AppendFormat("<td class='text-end'>{0}</td>", GetDbDecimal(row, "ExRate").ToString("N2"))

                ' 19. Actual PO Ref (ActualPORef)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "ActualPORef")))

                ' 20. Sub-Status (Status - เช่น Cancelled)
                Dim subStatus As String = GetDbString(row, "Status")
                Dim statusClass As String = ""
                If subStatus.ToLower() = "cancelled" Then
                    statusClass = "class='status-cancelled'"
                End If
                sb.AppendFormat("<td {0}>{1}</td>", statusClass, HttpUtility.HtmlEncode(subStatus))

                ' 21. Sub-Status Date (StatusDate)
                Dim statusDate As String = ""
                If row.Table.Columns.Contains("StatusDate") AndAlso row("StatusDate") IsNot DBNull.Value Then
                    statusDate = Convert.ToDateTime(row("StatusDate")).ToString("dd/MM/yyyy HH:mm tt")
                End If
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(statusDate))

                ' 22. Remark (Remark)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "Remark")))

                ' 23. Action By (ActionBy)
                sb.AppendFormat("<td>{0}</td>", HttpUtility.HtmlEncode(GetDbString(row, "ActionBy")))

                ' 24. Action Button (Static)
                ' (คุณอาจจะต้องเพิ่ม RunNo หรือ ID ให้ปุ่มนี้ ถ้าปุ่ม Edit ต้องทำงาน)
                sb.Append("<td><button class=""btn btn-action""><i class=""bi bi-pencil""></i> Edit</button></td>")

                sb.Append("</tr>")
            Next

        End If
        Return sb.ToString()
    End Function

    ''' <summary>
    ''' (Helper) ดึงค่า String จาก DataRow โดยเช็ค DBNull
    ''' </summary>
    Private Function GetDbString(row As DataRow, columnName As String) As String
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            Return row(columnName).ToString().Trim()
        End If
        Return ""
    End Function

    ''' <summary>
    ''' (Helper) ดึงค่า Decimal จาก DataRow โดยเช็ค DBNull
    ''' </summary>
    Private Function GetDbDecimal(row As DataRow, columnName As String) As Decimal
        Dim val As Decimal = 0
        If row IsNot Nothing AndAlso row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            Decimal.TryParse(row(columnName).ToString(), val)
        End If
        Return val
    End Function

    ''' <summary>
    ''' (Helper) แปลงเลขเดือนเป็นชื่อย่อ (เช่น "6" -> "Jun")
    ''' </summary>
    Private Function GetMonthName(month As Object) As String
        If month Is DBNull.Value OrElse month Is Nothing Then Return ""

        Dim monthStr As String = month.ToString()
        Dim monthInt As Integer
        If Not Integer.TryParse(monthStr, monthInt) Then
            Return monthStr ' ถ้าเป็นชื่ออยู่แล้ว ก็ส่งกลับเลย
        End If

        Select Case monthInt
            Case 1 : Return "Jan"
            Case 2 : Return "Feb"
            Case 3 : Return "Mar"
            Case 4 : Return "Apr"
            Case 5 : Return "May"
            Case 6 : Return "Jun"
            Case 7 : Return "Jul"
            Case 8 : Return "Aug"
            Case 9 : Return "Sep"
            Case 10 : Return "Oct"
            Case 11 : Return "Nov"
            Case 12 : Return "Dec"
            Case Else : Return monthInt.ToString()
        End Select
    End Function

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class