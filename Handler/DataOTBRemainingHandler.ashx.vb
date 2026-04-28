Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports Newtonsoft.Json ' ต้องมี Newtonsoft.Json ในโปรเจกต์

Public Class DataOTBRemainingHandler
    Implements IHttpHandler

    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.ContentType = "application/json"
        context.Response.ContentEncoding = Encoding.UTF8

        Try
            Dim year As String = NormalizeFilter(context.Request.Form("OTByear"))
            Dim month As String = NormalizeFilter(context.Request.Form("OTBmonth"))
            Dim company As String = NormalizeFilter(context.Request.Form("OTBCompany"))
            Dim category As String = NormalizeFilter(context.Request.Form("OTBCategory"))
            Dim segment As String = NormalizeFilter(context.Request.Form("OTBSegment"))
            Dim brand As String = NormalizeFilter(context.Request.Form("OTBBrand"))
            Dim vendor As String = NormalizeFilter(context.Request.Form("OTBVendor"))

            If String.IsNullOrEmpty(year) Then
                Throw New Exception("All filter fields are required to view the report.")
            End If

            Dim dtDetail As DataTable = GetRemainingDetail(year, month, company, category, segment, brand, vendor)
            dtDetail.TableName = "detail"

            Dim dtOther As New DataTable("otherRemaining")

            Dim ds As New DataSet()
            ds.Tables.Add(dtDetail)
            ds.Tables.Add(dtOther)

            Dim jsonResult As String = JsonConvert.SerializeObject(ds, Formatting.None)
            context.Response.Write(jsonResult)

        Catch ex As Exception
            context.Response.StatusCode = 200
            ' ส่ง Error กลับไปเป็น JSON
            Dim errorResponse As New With {
                .success = False,
                .message = ex.Message
            }
            context.Response.Write(JsonConvert.SerializeObject(errorResponse))
        End Try
    End Sub

    Private Function NormalizeFilter(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return Nothing
        Return value.Trim()
    End Function

    Private Function GetRemainingDetail(year As String, month As String, company As String, category As String,
                                        segment As String, brand As String, vendor As String) As DataTable
        Dim dt As New DataTable()
        Dim actualSegmentExpression As String = "SUBSTRING(ISNULL(a.Segment_Code, ''), 2, CASE WHEN LEN(ISNULL(a.Segment_Code, '')) > 2 THEN LEN(a.Segment_Code) - 2 ELSE 0 END)"

        Dim query As String = "
            WITH BudgetRows AS (
                SELECT
                    CASE WHEN [Type] = 'Original' THEN ISNULL(Amount, 0) ELSE 0 END AS Original,
                    CASE WHEN [Type] = 'Revise' THEN ISNULL(RevisedDiff, 0) ELSE 0 END AS RevDiff,
                    CAST(0 AS decimal(18,2)) AS Extra,
                    CAST(0 AS decimal(18,2)) AS SwitchIn,
                    CAST(0 AS decimal(18,2)) AS BalanceIn,
                    CAST(0 AS decimal(18,2)) AS CarryIn,
                    CAST(0 AS decimal(18,2)) AS SwitchOut,
                    CAST(0 AS decimal(18,2)) AS BalanceOut,
                    CAST(0 AS decimal(18,2)) AS CarryOut
                FROM [BMS].[dbo].[OTB_Transaction]
                WHERE OTBStatus = 'Approved'
                  AND (@Year IS NULL OR [Year] = @Year)
                  AND (@Month IS NULL OR [Month] = @Month)
                  AND (@Company IS NULL OR Company = @Company)
                  AND (@Category IS NULL OR Category = @Category)
                  AND (@Segment IS NULL OR Segment = @Segment)
                  AND (@Brand IS NULL OR Brand = @Brand)
                  AND (@Vendor IS NULL OR Vendor = @Vendor)

                UNION ALL

                SELECT
                    CAST(0 AS decimal(18,2)) AS Original,
                    CAST(0 AS decimal(18,2)) AS RevDiff,
                    CASE WHEN [From] = 'E' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS Extra,
                    CAST(0 AS decimal(18,2)) AS SwitchIn,
                    CAST(0 AS decimal(18,2)) AS BalanceIn,
                    CAST(0 AS decimal(18,2)) AS CarryIn,
                    CASE WHEN [From] = 'D' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS SwitchOut,
                    CASE WHEN [From] = 'I' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS BalanceOut,
                    CASE WHEN [From] = 'G' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS CarryOut
                FROM [BMS].[dbo].[OTB_Switching_Transaction]
                WHERE OTBStatus = 'Approved'
                  AND (@Year IS NULL OR [Year] = @Year)
                  AND (@Month IS NULL OR [Month] = @Month)
                  AND (@Company IS NULL OR Company = @Company)
                  AND (@Category IS NULL OR Category = @Category)
                  AND (@Segment IS NULL OR Segment = @Segment)
                  AND (@Brand IS NULL OR Brand = @Brand)
                  AND (@Vendor IS NULL OR Vendor = @Vendor)

                UNION ALL

                SELECT
                    CAST(0 AS decimal(18,2)) AS Original,
                    CAST(0 AS decimal(18,2)) AS RevDiff,
                    CAST(0 AS decimal(18,2)) AS Extra,
                    CASE WHEN [To] = 'C' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS SwitchIn,
                    CASE WHEN [To] = 'H' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS BalanceIn,
                    CASE WHEN [To] = 'F' THEN ISNULL(BudgetAmount, 0) ELSE 0 END AS CarryIn,
                    CAST(0 AS decimal(18,2)) AS SwitchOut,
                    CAST(0 AS decimal(18,2)) AS BalanceOut,
                    CAST(0 AS decimal(18,2)) AS CarryOut
                FROM [BMS].[dbo].[OTB_Switching_Transaction]
                WHERE OTBStatus = 'Approved'
                  AND [To] IS NOT NULL
                  AND SwitchYear IS NOT NULL
                  AND SwitchMonth IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchCompany)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchCategory)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchSegment)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchBrand)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(SwitchVendor)), '') IS NOT NULL
                  AND (@Year IS NULL OR SwitchYear = @Year)
                  AND (@Month IS NULL OR SwitchMonth = @Month)
                  AND (@Company IS NULL OR SwitchCompany = @Company)
                  AND (@Category IS NULL OR SwitchCategory = @Category)
                  AND (@Segment IS NULL OR SwitchSegment = @Segment)
                  AND (@Brand IS NULL OR SwitchBrand = @Brand)
                  AND (@Vendor IS NULL OR SwitchVendor = @Vendor)
            ),
            BudgetTotals AS (
                SELECT
                    SUM(Original) AS Original,
                    SUM(RevDiff) AS RevDiff,
                    SUM(Extra) AS Extra,
                    SUM(SwitchIn) AS SwitchIn,
                    SUM(BalanceIn) AS BalanceIn,
                    SUM(CarryIn) AS CarryIn,
                    SUM(SwitchOut) AS SwitchOut,
                    SUM(BalanceOut) AS BalanceOut,
                    SUM(CarryOut) AS CarryOut
                FROM BudgetRows
            ),
            UsageRows AS (
                SELECT
                    SUM(ISNULL(d.Amount_THB, 0)) AS DraftPO,
                    CAST(0 AS decimal(18,2)) AS ActualPO
                FROM [BMS].[dbo].[Draft_PO_Transaction] d
                WHERE ISNULL(d.[Status], 'Draft') NOT IN ('Matched', 'ForceMatching', 'Matching', 'Cancelled')
                  AND NOT EXISTS (
                      SELECT 1
                      FROM [BMS].[dbo].[Actual_PO_Summary] a2
                      WHERE a2.[Status] = 'Matched'
                        AND (a2.Draft_PO_Ref = d.DraftPO_No OR a2.PO_No = d.Actual_PO_No)
                  )
                  AND (@Year IS NULL OR d.PO_Year = @Year)
                  AND (@Month IS NULL OR d.PO_Month = @Month)
                  AND (@Company IS NULL OR d.Company_Code = @Company)
                  AND (@Category IS NULL OR d.Category_Code = @Category)
                  AND (@Segment IS NULL OR d.Segment_Code = @Segment)
                  AND (@Brand IS NULL OR d.Brand_Code = @Brand)
                  AND (@Vendor IS NULL OR d.Vendor_Code = @Vendor)

                UNION ALL

                SELECT
                    CAST(0 AS decimal(18,2)) AS DraftPO,
                    SUM(ISNULL(a.Amount_THB, 0)) AS ActualPO
                FROM [BMS].[dbo].[Actual_PO_Summary] a
                WHERE ISNULL(a.[Status], '') IN ('Matching', 'ForceMatching', 'Matched')
                  AND a.OTB_Year IS NOT NULL
                  AND a.OTB_Month IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(a.Company_Code)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(a.Category_Code)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(" & actualSegmentExpression & ")), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(a.Brand_Code)), '') IS NOT NULL
                  AND NULLIF(LTRIM(RTRIM(a.Vendor_Code)), '') IS NOT NULL
                  AND (@Year IS NULL OR a.OTB_Year = @Year)
                  AND (@Month IS NULL OR a.OTB_Month = @Month)
                  AND (@Company IS NULL OR a.Company_Code = @Company)
                  AND (@Category IS NULL OR a.Category_Code = @Category)
                  AND (@Segment IS NULL OR " & actualSegmentExpression & " = @Segment)
                  AND (@Brand IS NULL OR a.Brand_Code = @Brand)
                  AND (@Vendor IS NULL OR a.Vendor_Code = @Vendor)
            ),
            UsageTotals AS (
                SELECT
                    SUM(DraftPO) AS DraftPO,
                    SUM(ActualPO) AS ActualPO
                FROM UsageRows
            )
            SELECT
                ISNULL(b.Original, 0) AS BudgetApproved_Original,
                ISNULL(b.RevDiff, 0) AS RevisedDiff,
                ISNULL(b.Extra, 0) AS Budget_Extra,
                ISNULL(b.SwitchIn, 0) AS Budget_SwitchIn,
                ISNULL(b.BalanceIn, 0) AS Budget_BalanceIn,
                ISNULL(b.CarryIn, 0) AS Budget_CarryIn,
                ISNULL(b.SwitchOut, 0) AS Budget_SwitchOut,
                ISNULL(b.BalanceOut, 0) AS Budget_BalanceOut,
                ISNULL(b.CarryOut, 0) AS Budget_CarryOut,
                ISNULL(b.Original, 0) + ISNULL(b.RevDiff, 0) + ISNULL(b.Extra, 0)
                    + ISNULL(b.SwitchIn, 0) + ISNULL(b.BalanceIn, 0) + ISNULL(b.CarryIn, 0)
                    - ISNULL(b.SwitchOut, 0) - ISNULL(b.BalanceOut, 0) - ISNULL(b.CarryOut, 0) AS TotalBudgetApproved,
                ISNULL(u.DraftPO, 0) AS TotalDraftPO,
                ISNULL(u.ActualPO, 0) AS TotalActualPO,
                ISNULL(u.DraftPO, 0) + ISNULL(u.ActualPO, 0) AS TotalPO_Usage,
                (ISNULL(b.Original, 0) + ISNULL(b.RevDiff, 0) + ISNULL(b.Extra, 0)
                    + ISNULL(b.SwitchIn, 0) + ISNULL(b.BalanceIn, 0) + ISNULL(b.CarryIn, 0)
                    - ISNULL(b.SwitchOut, 0) - ISNULL(b.BalanceOut, 0) - ISNULL(b.CarryOut, 0))
                    - (ISNULL(u.DraftPO, 0) + ISNULL(u.ActualPO, 0)) AS Remaining
            FROM BudgetTotals b
            CROSS JOIN UsageTotals u"

            Using conn As New SqlConnection(connectionString)
                conn.Open()
                Using cmd As New SqlCommand(query, conn)
                    AddFilterParameters(cmd, year, month, company, category, segment, brand, vendor)
                    Using adapter As New SqlDataAdapter(cmd)
                        adapter.Fill(dt)
                    End Using
                End Using
            End Using

            Return dt
    End Function

    Private Sub AddFilterParameters(cmd As SqlCommand, year As String, month As String, company As String,
                                    category As String, segment As String, brand As String, vendor As String)
        AddOptionalIntParameter(cmd, "@Year", year, "Year")
        AddOptionalIntParameter(cmd, "@Month", month, "Month")
        AddOptionalNVarCharParameter(cmd, "@Company", company, 20)
        AddOptionalNVarCharParameter(cmd, "@Category", category, 20)
        AddOptionalNVarCharParameter(cmd, "@Segment", segment, 20)
        AddOptionalNVarCharParameter(cmd, "@Brand", brand, 30)
        AddOptionalNVarCharParameter(cmd, "@Vendor", vendor, 30)
    End Sub

    Private Sub AddOptionalIntParameter(cmd As SqlCommand, name As String, value As String, label As String)
        Dim parsedValue As Integer
        Dim parameter As SqlParameter = cmd.Parameters.Add(name, SqlDbType.Int)
        If String.IsNullOrWhiteSpace(value) Then
            parameter.Value = DBNull.Value
        ElseIf Integer.TryParse(value, parsedValue) Then
            parameter.Value = parsedValue
        Else
            Throw New Exception(label & " is invalid.")
        End If
    End Sub

    Private Sub AddOptionalNVarCharParameter(cmd As SqlCommand, name As String, value As String, size As Integer)
        Dim parameter As SqlParameter = cmd.Parameters.Add(name, SqlDbType.NVarChar, size)
        If String.IsNullOrWhiteSpace(value) Then
            parameter.Value = DBNull.Value
        Else
            parameter.Value = value.Trim()
        End If
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class
