Imports System.Data
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Linq

Public Class OTBSwitchBudgetGuard
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
        Return String.Join("|", New String() {year, month, company, category, segment, brand, vendor})
    End Function

    Public Shared Sub AcquireSourceLocks(connection As SqlConnection, sourceKeys As IEnumerable(Of String))
        If connection Is Nothing OrElse connection.State <> ConnectionState.Open Then
            Throw New InvalidOperationException("An open database connection is required for the OTB budget lock.")
        End If

        Dim keys = sourceKeys.Where(Function(k) Not String.IsNullOrWhiteSpace(k)).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(Function(k) k, StringComparer.OrdinalIgnoreCase)
        For Each key In keys
            Using cmd As New SqlCommand("DECLARE @Result int; EXEC @Result = sys.sp_getapplock @Resource=@Resource, @LockMode='Exclusive', @LockOwner='Session', @LockTimeout=30000; SELECT @Result;", connection)
                cmd.Parameters.Add("@Resource", SqlDbType.NVarChar, 255).Value = "BMS_OTB_SWITCH|" & key
                Dim lockResult = Convert.ToInt32(cmd.ExecuteScalar())
                If lockResult < 0 Then
                    Throw New InvalidOperationException("Unable to lock the OTB source for validation. Please try again.")
                End If
            End Using
        Next
    End Sub
End Class
