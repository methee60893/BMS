Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Linq
Imports System.Globalization

Public Class OTBSwitchBudgetGuard
    Private NotInheritable Class BudgetLockHandle
        Implements IDisposable

        Private ReadOnly _connection As SqlConnection
        Private ReadOnly _resources As List(Of String)
        Private _disposed As Boolean

        Public Sub New(connection As SqlConnection, resources As IEnumerable(Of String))
            _connection = connection
            _resources = resources.ToList()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            ReleaseBudgetLocks(_connection, _resources)
        End Sub
    End Class

    Public Class BudgetDimension
        Public Property Year As Integer
        Public Property Month As Integer
        Public Property Company As String
        Public Property Category As String
        Public Property Segment As String
        Public Property Brand As String
        Public Property Vendor As String
    End Class

    Public Class BudgetCheckResult
        Public Property ApprovedBudget As Decimal
        Public Property DraftPO As Decimal
        Public Property ActualPO As Decimal
        Public Property AvailableBudget As Decimal
    End Class

    Public Shared Function Check(year As String, month As String, category As String, company As String,
                                 segment As String, brand As String, vendor As String,
                                 Optional calculator As OTBBudgetCalculator = Nothing,
                                 Optional poValidator As POValidate = Nothing) As BudgetCheckResult
        If calculator Is Nothing Then calculator = New OTBBudgetCalculator()
        If poValidator Is Nothing Then poValidator = New POValidate()

        Dim approved = calculator.CalculateCurrentApprovedBudget(year, month, category, company, segment, brand, vendor)
        Dim usage = poValidator.GetUsedBudgetBreakdownForOTB(year, month, category, company, segment, brand, vendor)

        Return New BudgetCheckResult With {
            .ApprovedBudget = approved,
            .DraftPO = usage("DraftPO"),
            .ActualPO = usage("ActualPO"),
            .AvailableBudget = approved - usage("DraftPO") - usage("ActualPO")
        }
    End Function

    Public Shared Sub EnsureSufficient(result As BudgetCheckResult, requestedAmount As Decimal)
        If requestedAmount <= 0 Then
            Throw New InvalidOperationException("Switch amount must be greater than 0.")
        End If

        If requestedAmount > result.AvailableBudget Then
            Throw New InvalidOperationException(
                $"Insufficient OTB Remaining. Approved: {result.ApprovedBudget:N2} THB, " &
                $"Draft PO: {result.DraftPO:N2} THB, Actual PO (Matched): {result.ActualPO:N2} THB, " &
                $"Available: {result.AvailableBudget:N2} THB, Requested: {requestedAmount:N2} THB.")
        End If
    End Sub

    Public Shared Function BuildSourceKey(year As String, month As String, company As String, category As String,
                                          segment As String, brand As String, vendor As String) As String
        Dim dimension As BudgetDimension = CreateDimension(year, month, company, category, segment, brand, vendor)
        Return String.Join("|", New String() {
            dimension.Year.ToString(CultureInfo.InvariantCulture),
            dimension.Month.ToString(CultureInfo.InvariantCulture),
            dimension.Company, dimension.Category, dimension.Segment, dimension.Brand, dimension.Vendor})
    End Function

    Public Shared Function CreateDimension(year As String, month As String, company As String, category As String,
                                           segment As String, brand As String, vendor As String) As BudgetDimension
        Dim yearValue As Integer
        Dim monthValue As Integer
        If Not Integer.TryParse(If(year, "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, yearValue) Then
            Throw New InvalidOperationException("OTB year must be numeric before acquiring a budget lock.")
        End If
        If Not Integer.TryParse(If(month, "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, monthValue) OrElse monthValue < 1 OrElse monthValue > 12 Then
            Throw New InvalidOperationException("OTB month must be between 1 and 12 before acquiring a budget lock.")
        End If
        Return New BudgetDimension With {
            .Year = yearValue,
            .Month = monthValue,
            .Company = NormalizeRequiredCode(company, "Company"),
            .Category = NormalizeRequiredCode(category, "Category"),
            .Segment = NormalizeRequiredCode(segment, "Segment"),
            .Brand = NormalizeRequiredCode(brand, "Brand"),
            .Vendor = NormalizeRequiredCode(vendor, "Vendor")
        }
    End Function

    Public Shared Function BuildGroupLockResource(dimension As BudgetDimension) As String
        If dimension Is Nothing Then Throw New ArgumentNullException(NameOf(dimension))
        Return String.Join("|", New String() {
            "BMS_OTB_GROUP", dimension.Company,
            dimension.Year.ToString(CultureInfo.InvariantCulture),
            dimension.Month.ToString(CultureInfo.InvariantCulture), dimension.Category})
    End Function

    Public Shared Function BuildDetailLockResource(dimension As BudgetDimension) As String
        If dimension Is Nothing Then Throw New ArgumentNullException(NameOf(dimension))
        Return String.Join("|", New String() {
            "BMS_OTB_DETAIL", dimension.Year.ToString(CultureInfo.InvariantCulture),
            dimension.Month.ToString(CultureInfo.InvariantCulture), dimension.Company, dimension.Category,
            dimension.Segment, dimension.Brand, dimension.Vendor})
    End Function

    Public Shared Function BuildGroupLockResources(dimensions As IEnumerable(Of BudgetDimension)) As List(Of String)
        If dimensions Is Nothing Then Return New List(Of String)()
        Return dimensions.Where(Function(item) item IsNot Nothing).
            Select(Function(item) BuildGroupLockResource(item)).
            Distinct(StringComparer.Ordinal).
            OrderBy(Function(item) item, StringComparer.Ordinal).
            ToList()
    End Function

    Public Shared Function AcquireBudgetLocks(connection As SqlConnection, resources As IEnumerable(Of String)) As IDisposable
        If connection Is Nothing OrElse connection.State <> ConnectionState.Open Then
            Throw New InvalidOperationException("An open database connection is required for the OTB budget lock.")
        End If

        Dim orderedResources As List(Of String) = If(resources, Enumerable.Empty(Of String)()).
            Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
            Select(Function(item) item.Trim().ToUpperInvariant()).
            Distinct(StringComparer.Ordinal).
            OrderBy(Function(item) item, StringComparer.Ordinal).
            ToList()
        Dim acquired As New List(Of String)()
        Try
            For Each resource As String In orderedResources
                Using cmd As New SqlCommand("DECLARE @Result int; EXEC @Result = sys.sp_getapplock @Resource=@Resource, @LockMode='Exclusive', @LockOwner='Session', @LockTimeout=30000; SELECT @Result;", connection)
                    cmd.Parameters.Add("@Resource", SqlDbType.NVarChar, 255).Value = resource
                    Dim lockResult = Convert.ToInt32(cmd.ExecuteScalar())
                    If lockResult < 0 Then
                        Throw New InvalidOperationException("Unable to lock the OTB budget group for validation. Please try again.")
                    End If
                    acquired.Add(resource)
                End Using
            Next
            Return New BudgetLockHandle(connection, acquired)
        Catch
            ReleaseBudgetLocks(connection, acquired)
            Throw
        End Try
    End Function

    Public Shared Function AcquireSourceLocks(connection As SqlConnection, sourceKeys As IEnumerable(Of String)) As IDisposable
        Dim resources As IEnumerable(Of String) = If(sourceKeys, Enumerable.Empty(Of String)()).
            Where(Function(item) Not String.IsNullOrWhiteSpace(item)).
            Select(Function(item) "BMS_OTB_DETAIL|" & item.Trim().ToUpperInvariant())
        Return AcquireBudgetLocks(connection, resources)
    End Function

    ''' <summary>
    ''' Revalidates every submitted movement dimension against live master data.
    ''' Preview JSON and manual form values are client-controlled and must not be
    ''' forwarded to SAP merely because their fields are non-empty.
    ''' </summary>
    Public Shared Sub EnsureValidMasterDimensions(connection As SqlConnection, dimensions As IEnumerable(Of BudgetDimension))
        If connection Is Nothing OrElse connection.State <> ConnectionState.Open Then
            Throw New InvalidOperationException("An open database connection is required to validate OTB dimensions.")
        End If
        Dim unique As List(Of BudgetDimension) = If(dimensions, Enumerable.Empty(Of BudgetDimension)()).
            Where(Function(item) item IsNot Nothing).
            GroupBy(Function(item) BuildDetailLockResource(item), StringComparer.Ordinal).
            Select(Function(group) group.First()).ToList()
        If unique.Count = 0 Then Throw New InvalidOperationException("At least one OTB budget dimension is required.")

        Dim stage As New DataTable()
        stage.Columns.Add("Year", GetType(Integer))
        stage.Columns.Add("Month", GetType(Integer))
        stage.Columns.Add("Company", GetType(String))
        stage.Columns.Add("Category", GetType(String))
        stage.Columns.Add("Segment", GetType(String))
        stage.Columns.Add("Brand", GetType(String))
        stage.Columns.Add("Vendor", GetType(String))
        For Each item As BudgetDimension In unique
            stage.Rows.Add(item.Year, item.Month, item.Company, item.Category, item.Segment, item.Brand, item.Vendor)
        Next

        Using createCmd As New SqlCommand("
            CREATE TABLE #OTBMasterDimensions
            (
                [Year] int NOT NULL,
                [Month] int NOT NULL,
                Company nvarchar(20) NOT NULL,
                Category nvarchar(20) NOT NULL,
                Segment nvarchar(20) NOT NULL,
                Brand nvarchar(30) NOT NULL,
                Vendor nvarchar(30) NOT NULL,
                PRIMARY KEY ([Year], [Month], Company, Category, Segment, Brand, Vendor)
            );", connection)
            createCmd.CommandTimeout = 300
            createCmd.ExecuteNonQuery()
        End Using
        Using bulk As New SqlBulkCopy(connection)
            bulk.DestinationTableName = "#OTBMasterDimensions"
            bulk.BatchSize = Math.Min(stage.Rows.Count, 2000)
            bulk.BulkCopyTimeout = 300
            For Each column As DataColumn In stage.Columns
                bulk.ColumnMappings.Add(column.ColumnName, column.ColumnName)
            Next
            bulk.WriteToServer(stage)
        End Using
        Using cmd As New SqlCommand("
            SELECT TOP (1) d.[Year]
            FROM #OTBMasterDimensions d
            WHERE NOT EXISTS (SELECT 1 FROM dbo.MS_Year y WHERE y.Year_Kept = d.[Year])
               OR NOT EXISTS (SELECT 1 FROM dbo.MS_Month m WHERE m.month_code = d.[Month])
               OR NOT EXISTS (SELECT 1 FROM dbo.MS_Company c WHERE c.CompanyCode = d.Company)
               OR NOT EXISTS (SELECT 1 FROM dbo.MS_Category c WHERE c.Cate = d.Category)
               OR NOT EXISTS (SELECT 1 FROM dbo.MS_Segment s WHERE s.SegmentCode = d.Segment)
               OR NOT EXISTS (SELECT 1 FROM dbo.MS_Brand b WHERE b.[Brand Code] = d.Brand)
               OR NOT EXISTS (SELECT 1 FROM dbo.MS_Vendor v WHERE v.VendorCode = d.Vendor);", connection)
            cmd.CommandTimeout = 300
            Dim invalid As Object = cmd.ExecuteScalar()
            If invalid IsNot Nothing AndAlso invalid IsNot DBNull.Value Then
                Throw New InvalidOperationException("One or more OTB dimensions are not present in live master data. SAP was not called.")
            End If
        End Using
    End Sub

    Public Shared Sub EnsureNoActiveApprovalClaims(connection As SqlConnection, dimensions As IEnumerable(Of BudgetDimension))
        If connection Is Nothing OrElse connection.State <> ConnectionState.Open Then
            Throw New InvalidOperationException("An open database connection is required to check approval claims.")
        End If
        Dim unique As List(Of BudgetDimension) = If(dimensions, Enumerable.Empty(Of BudgetDimension)()).
            Where(Function(item) item IsNot Nothing).
            GroupBy(Function(item) BuildGroupLockResource(item), StringComparer.Ordinal).
            Select(Function(group) group.First()).ToList()
        If unique.Count = 0 Then Return

        Dim stage As New DataTable()
        stage.Columns.Add("Company", GetType(String))
        stage.Columns.Add("Year", GetType(Integer))
        stage.Columns.Add("Month", GetType(Integer))
        stage.Columns.Add("Category", GetType(String))
        For Each item As BudgetDimension In unique
            stage.Rows.Add(item.Company, item.Year, item.Month, item.Category)
        Next

        Using createCmd As New SqlCommand("
            CREATE TABLE #OTBClaimGroups
            (
                Company nvarchar(20) NOT NULL,
                [Year] int NOT NULL,
                [Month] int NOT NULL,
                Category nvarchar(20) NOT NULL,
                PRIMARY KEY (Company, [Year], [Month], Category)
            );", connection)
            createCmd.ExecuteNonQuery()
        End Using
        Using bulk As New SqlBulkCopy(connection)
            bulk.DestinationTableName = "#OTBClaimGroups"
            bulk.BatchSize = Math.Min(stage.Rows.Count, 2000)
            bulk.BulkCopyTimeout = 300
            For Each column As DataColumn In stage.Columns
                bulk.ColumnMappings.Add(column.ColumnName, column.ColumnName)
            Next
            bulk.WriteToServer(stage)
        End Using
        Using cmd As New SqlCommand("
            SELECT TOP (1) c.RunNo
            FROM dbo.Draft_OTB_Approval_Claim c WITH (INDEX(IX_Draft_OTB_Approval_Claim_Group))
            INNER JOIN #OTBClaimGroups g
                    ON g.Company = c.Company AND g.[Year] = c.[Year]
                   AND g.[Month] = c.[Month] AND g.Category = c.Category;", connection)
            cmd.CommandTimeout = 300
            Dim conflict As Object = cmd.ExecuteScalar()
            If conflict IsNot Nothing AndAlso conflict IsNot DBNull.Value Then
                Throw New InvalidOperationException("An approval or reconciliation job is active for one or more selected budget groups. SAP was not called.")
            End If
        End Using
    End Sub

    Private Shared Function NormalizeRequiredCode(value As String, fieldName As String) As String
        Dim normalized As String = If(value, "").Trim().ToUpperInvariant()
        If normalized.Length = 0 Then Throw New InvalidOperationException(fieldName & " is required before acquiring a budget lock.")
        Return normalized
    End Function

    Private Shared Sub ReleaseBudgetLocks(connection As SqlConnection, resources As IEnumerable(Of String))
        If connection Is Nothing OrElse connection.State <> ConnectionState.Open OrElse resources Is Nothing Then Return
        For Each resource As String In resources.Reverse()
            Try
                Using cmd As New SqlCommand("DECLARE @Result int; EXEC @Result = sys.sp_releaseapplock @Resource=@Resource, @LockOwner='Session'; SELECT @Result;", connection)
                    cmd.Parameters.Add("@Resource", SqlDbType.NVarChar, 255).Value = resource
                    cmd.ExecuteScalar()
                End Using
            Catch
                ' The connection is about to be disposed; release is best effort if SQL is already unavailable.
            End Try
        Next
    End Sub
End Class
