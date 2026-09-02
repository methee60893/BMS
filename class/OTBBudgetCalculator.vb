Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization

Public Class OTBBudgetCalculator
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    ' --- 1. สร้างตัวแปร Private สำหรับเก็บข้อมูลที่โหลดมา ---
    Private dtOtbTransaction As DataTable
    Private dtSwitchingTransaction As DataTable

    ' Immutable aggregate indexes turn each budget lookup into O(1), rather than
    ' rescanning all approved OTB and switching rows for every uploaded record.
    Private otbAmountsByKey As Dictionary(Of String, BudgetAggregate)
    Private switchingAmountsByKey As Dictionary(Of String, SwitchingAggregate)

    Private NotInheritable Class BudgetAggregate
        Public Original As Decimal
        Public RevisedDiff As Decimal
    End Class

    Private NotInheritable Class SwitchingAggregate
        Public Extra As Decimal
        Public SwitchIn As Decimal
        Public SwitchOut As Decimal
        Public BalanceIn As Decimal
        Public BalanceOut As Decimal
        Public CarryIn As Decimal
        Public CarryOut As Decimal
    End Class

    ''' <summary>
    ''' Constructor: เมื่อคลาสนี้ถูก New() จะดึงข้อมูลทั้งหมดจาก DB มาเก็บไว้ก่อน
    ''' </summary>
    Public Sub New()
        LoadAllTransactionData()
        BuildIndexes()
    End Sub

    Private Sub BuildIndexes()
        otbAmountsByKey = New Dictionary(Of String, BudgetAggregate)(StringComparer.OrdinalIgnoreCase)
        switchingAmountsByKey = New Dictionary(Of String, SwitchingAggregate)(StringComparer.OrdinalIgnoreCase)

        If dtOtbTransaction IsNot Nothing Then
            For Each row As DataRow In dtOtbTransaction.Rows
                Dim key As String = BuildKey(RowText(row, "Year"), RowText(row, "Month"),
                                             RowText(row, "Category"), RowText(row, "Company"),
                                             RowText(row, "Segment"), RowText(row, "Brand"),
                                             RowText(row, "Vendor"))
                Dim aggregate As BudgetAggregate = Nothing
                If Not otbAmountsByKey.TryGetValue(key, aggregate) Then
                    aggregate = New BudgetAggregate()
                    otbAmountsByKey.Add(key, aggregate)
                End If

                Dim transactionType As String = NormalizeTextComponent(RowText(row, "Type"))
                If transactionType.Equals("Original", StringComparison.OrdinalIgnoreCase) Then
                    AddRowAmount(row, "Amount", aggregate.Original)
                ElseIf transactionType.Equals("Revise", StringComparison.OrdinalIgnoreCase) Then
                    AddRowAmount(row, "RevisedDiff", aggregate.RevisedDiff)
                End If
            Next
        End If

        If dtSwitchingTransaction IsNot Nothing Then
            For Each row As DataRow In dtSwitchingTransaction.Rows
                Dim amount As Decimal = 0
                If row.Table.Columns.Contains("BudgetAmount") AndAlso row("BudgetAmount") IsNot DBNull.Value Then
                    amount = Convert.ToDecimal(row("BudgetAmount"), CultureInfo.InvariantCulture)
                End If

                Dim source As SwitchingAggregate = GetOrCreateSwitchingAggregate(
                    BuildKey(RowText(row, "Year"), RowText(row, "Month"),
                             RowText(row, "Category"), RowText(row, "Company"),
                             RowText(row, "Segment"), RowText(row, "Brand"),
                             RowText(row, "Vendor")))
                Select Case RowText(row, "From").Trim().ToUpperInvariant()
                    Case "D" : source.SwitchOut += amount
                    Case "G" : source.CarryOut += amount
                    Case "I" : source.BalanceOut += amount
                    Case "E" : source.Extra += amount
                End Select

                ' The old destination filter explicitly required To IS NOT NULL.
                If row.Table.Columns.Contains("To") AndAlso row("To") IsNot DBNull.Value Then
                    Dim destination As SwitchingAggregate = GetOrCreateSwitchingAggregate(
                        BuildKey(RowText(row, "SwitchYear"), RowText(row, "SwitchMonth"),
                                 RowText(row, "SwitchCategory"), RowText(row, "SwitchCompany"),
                                 RowText(row, "SwitchSegment"), RowText(row, "SwitchBrand"),
                                 RowText(row, "SwitchVendor")))
                    Select Case RowText(row, "To").Trim().ToUpperInvariant()
                        Case "C" : destination.SwitchIn += amount
                        Case "F" : destination.CarryIn += amount
                        Case "H" : destination.BalanceIn += amount
                    End Select
                End If
            Next
        End If
    End Sub

    Private Function GetOrCreateSwitchingAggregate(key As String) As SwitchingAggregate
        Dim aggregate As SwitchingAggregate = Nothing
        If Not switchingAmountsByKey.TryGetValue(key, aggregate) Then
            aggregate = New SwitchingAggregate()
            switchingAmountsByKey.Add(key, aggregate)
        End If
        Return aggregate
    End Function

    Private Shared Sub AddRowAmount(row As DataRow, columnName As String, ByRef total As Decimal)
        If row.Table.Columns.Contains(columnName) AndAlso row(columnName) IsNot DBNull.Value Then
            total += Convert.ToDecimal(row(columnName), CultureInfo.InvariantCulture)
        End If
    End Sub

    Private Shared Function RowText(row As DataRow, columnName As String) As String
        If row Is Nothing OrElse Not row.Table.Columns.Contains(columnName) OrElse row(columnName) Is DBNull.Value Then
            Return ""
        End If
        Return Convert.ToString(row(columnName), CultureInfo.InvariantCulture)
    End Function

    Private Shared Function NormalizeNumericComponent(value As String) As String
        Dim parsed As Integer
        Dim raw As String = If(value, "")
        If Integer.TryParse(raw.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
            Return parsed.ToString(CultureInfo.InvariantCulture)
        End If
        Return raw
    End Function

    Private Shared Function NormalizeTextComponent(value As String) As String
        Return If(value, "").TrimEnd(" "c)
    End Function

    Private Shared Function BuildCompositeKey(ParamArray values() As String) As String
        Dim key As New StringBuilder()
        For Each value As String In values
            Dim component As String = NormalizeTextComponent(value)
            key.Append(component.Length.ToString(CultureInfo.InvariantCulture)).Append(":"c).Append(component).Append("|"c)
        Next
        Return key.ToString()
    End Function

    Private Shared Function BuildKey(year As String, month As String, category As String,
                                     company As String, segment As String, brand As String,
                                     vendor As String) As String
        Return BuildCompositeKey(NormalizeNumericComponent(year), NormalizeNumericComponent(month),
                                 category, company, segment, brand, vendor)
    End Function

    ''' <summary>
    ''' (ใหม่) เปิด Connection ครั้งเดียว แล้วดึงข้อมูลทั้งหมดที่ Approved แล้วมาเก็บไว้
    ''' </summary>
    Private Sub LoadAllTransactionData()
        dtOtbTransaction = New DataTable()
        dtSwitchingTransaction = New DataTable()

        Try
            Using conn As New SqlConnection(connectionString)
                conn.Open()

                ' --- โหลดข้อมูล OTB_Transaction (สำหรับ Original และ Revise) ---
                Dim queryOtb As String = "
                    SELECT [Type], [Year], [Month], [Category], [Company], [Segment], 
                           [Brand], [Vendor], [Amount], [RevisedDiff], [Version]
                    FROM [dbo].[OTB_Transaction] 
                    WHERE [OTBStatus] = 'Approved'"

                Using cmd As New SqlCommand(queryOtb, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtOtbTransaction.Load(reader)
                    End Using
                End Using

                ' --- โหลดข้อมูล OTB_Switching_Transaction (สำหรับ Extra, Switch, Balance, Carry) ---
                Dim querySwitch As String = "
                    SELECT [Year], [Month], [Category], [Company], [Segment], [Brand], [Vendor], 
                           [From], [BudgetAmount], 
                           [To], [SwitchYear], [SwitchMonth], [SwitchCompany], [SwitchCategory], 
                           [SwitchSegment], [SwitchBrand], [SwitchVendor] 
                    FROM [dbo].[OTB_Switching_Transaction] 
                    WHERE [OTBStatus] = 'Approved'"

                Using cmd As New SqlCommand(querySwitch, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        dtSwitchingTransaction.Load(reader)
                    End Using
                End Using
            End Using
        Catch ex As Exception
            Throw New Exception("Failed to load budget calculator data: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' (แก้ไข) เปลี่ยนเป็น Instance Function (ลบ Shared)
    ''' </summary>
    Public Function CalculateCurrentApprovedBudget(year As String, month As String,
                                                   category As String, company As String,
                                                   segment As String, brand As String,
                                                   vendor As String) As Decimal
        Dim totalBudget As Decimal = 0
        Try
            totalBudget += GetOriginalAmount(year, month, category, company, segment, brand, vendor)
            totalBudget += GetRevisionDiff(year, month, category, company, segment, brand, vendor)

            Dim switchingAmounts As Dictionary(Of String, Decimal) = CalculateSwitchingAmounts(year, month, category, company, segment, brand, vendor)

            totalBudget += switchingAmounts("Extra")
            totalBudget += switchingAmounts("SwitchIn")
            totalBudget += switchingAmounts("BalanceIn")
            totalBudget += switchingAmounts("CarryIn")
            totalBudget -= switchingAmounts("SwitchOut")
            totalBudget -= switchingAmounts("BalanceOut")
            totalBudget -= switchingAmounts("CarryOut")

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Error calculating budget for key [{year}-{month}-{category}...]: {ex.Message}")
            Return 0
        End Try
        Return totalBudget
    End Function

    ''' <summary>
    ''' (แก้ไข) ลบ Shared และ 'conn' Parameter, เปลี่ยนไปใช้ .Compute จาก dtOtbTransaction
    ''' </summary>
    Private Function GetOriginalAmount(year As String, month As String,
                                       category As String, company As String, segment As String,
                                       brand As String, vendor As String) As Decimal
        Try
            Dim aggregate As BudgetAggregate = Nothing
            If otbAmountsByKey.TryGetValue(BuildKey(year, month, category, company, segment, brand, vendor), aggregate) Then
                Return aggregate.Original
            End If
            Return 0

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in GetOriginalAmount (In-Memory): " & ex.Message)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' (แก้ไข) ลบ Shared และ 'conn' Parameter, เปลี่ยนไปใช้ .Select จาก dtOtbTransaction
    ''' </summary>
    Private Function GetRevisionDiff(year As String, month As String,
                                     category As String, company As String, segment As String,
                                     brand As String, vendor As String) As Decimal
        Try
            Dim aggregate As BudgetAggregate = Nothing
            If otbAmountsByKey.TryGetValue(BuildKey(year, month, category, company, segment, brand, vendor), aggregate) Then
                Return aggregate.RevisedDiff
            End If
            Return 0

        Catch ex As Exception
            'System.Diagnostics.Debug.WriteLine("Error in GetRevisionDiff (In-Memory): " & ex.Message)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' (แก้ไข) ลบ Shared และ 'conn' Parameter, เปลี่ยนไปใช้ .Select จาก dtSwitchingTransaction
    ''' </summary>
    Private Function CalculateSwitchingAmounts(year As String, month As String,
                                               category As String, company As String, segment As String,
                                               brand As String, vendor As String) As Dictionary(Of String, Decimal)
        Dim result As New Dictionary(Of String, Decimal) From {
            {"Extra", 0}, {"SwitchIn", 0}, {"SwitchOut", 0},
            {"BalanceIn", 0}, {"BalanceOut", 0}, {"CarryIn", 0}, {"CarryOut", 0}
        }

        Try
            Dim aggregate As SwitchingAggregate = Nothing
            If switchingAmountsByKey.TryGetValue(BuildKey(year, month, category, company, segment, brand, vendor), aggregate) Then
                result("Extra") = aggregate.Extra
                result("SwitchIn") = aggregate.SwitchIn
                result("SwitchOut") = aggregate.SwitchOut
                result("BalanceIn") = aggregate.BalanceIn
                result("BalanceOut") = aggregate.BalanceOut
                result("CarryIn") = aggregate.CarryIn
                result("CarryOut") = aggregate.CarryOut
            End If

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CalculateSwitchingAmounts (In-Memory): " & ex.Message)
        End Try

        Return result
    End Function

    ''' <summary>
    ''' (แก้ไข) ดึงรายละเอียดการคำนวณ (สำหรับแสดง Breakdown) - ลบ Shared
    ''' </summary>
    Public Function GetBudgetBreakdown(year As String, month As String,
                                         category As String, company As String,
                                         segment As String, brand As String,
                                         vendor As String) As Dictionary(Of String, Decimal)
        Dim breakdown As New Dictionary(Of String, Decimal)

        Try
            ' ดึงข้อมูลแต่ละส่วน
            Dim originalAmt As Decimal = GetOriginalAmount(year, month, category, company, segment, brand, vendor)
            Dim revDiffAmt As Decimal = GetRevisionDiff(year, month, category, company, segment, brand, vendor)
            Dim switching As Dictionary(Of String, Decimal) = CalculateSwitchingAmounts(year, month, category, company, segment, brand, vendor)

            ' เพิ่มเข้า breakdown
            breakdown.Add("Original", originalAmt)
            breakdown.Add("RevDiff", revDiffAmt)
            breakdown.Add("Extra", switching("Extra"))
            breakdown.Add("SwitchIn", switching("SwitchIn"))
            breakdown.Add("BalanceIn", switching("BalanceIn"))
            breakdown.Add("CarryIn", switching("CarryIn"))
            breakdown.Add("SwitchOut", switching("SwitchOut"))
            breakdown.Add("BalanceOut", switching("BalanceOut"))
            breakdown.Add("CarryOut", switching("CarryOut"))

            ' คำนวณ Total
            Dim total As Decimal = originalAmt + revDiffAmt +
                                   switching("Extra") + switching("SwitchIn") + switching("BalanceIn") + switching("CarryIn") -
                                   switching("SwitchOut") - switching("BalanceOut") - switching("CarryOut")
            breakdown.Add("Total", total)

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error getting budget breakdown: " & ex.Message)
            ' (ควรจะ Clear และ Add default values 0 ถ้าต้องการให้ปลอดภัย)
            breakdown.Clear()
            breakdown.Add("Original", 0)
            breakdown.Add("RevDiff", 0)
            breakdown.Add("Extra", 0)
            breakdown.Add("SwitchIn", 0)
            breakdown.Add("BalanceIn", 0)
            breakdown.Add("CarryIn", 0)
            breakdown.Add("SwitchOut", 0)
            breakdown.Add("BalanceOut", 0)
            breakdown.Add("CarryOut", 0)
            breakdown.Add("Total", 0)
        End Try

        Return breakdown
    End Function

    ''' <summary>
    ''' สร้าง HTML Tooltip สำหรับแสดง Breakdown (ฟังก์ชันนี้เป็น Shared ได้ เพราะไม่ขึ้นกับข้อมูลใน Class)
    ''' </summary>
    Public Shared Function GetBreakdownTooltip(breakdown As Dictionary(Of String, Decimal)) As String
        Try
            Dim sb As New StringBuilder()

            sb.AppendLine("Budget Breakdown:")
            sb.AppendLine("─────────────────────")
            sb.AppendFormat("Original: {0:N2}{1}", breakdown("Original"), vbCrLf)
            sb.AppendFormat("Rev.Diff: {0:N2}{1}", breakdown("RevDiff"), vbCrLf)
            sb.AppendLine("─────────────────────")
            sb.AppendFormat("+ Extra: {0:N2}{1}", breakdown("Extra"), vbCrLf)
            sb.AppendFormat("+ Switch In: {0:N2}{1}", breakdown("SwitchIn"), vbCrLf)
            sb.AppendFormat("+ Balance In: {0:N2}{1}", breakdown("BalanceIn"), vbCrLf)
            sb.AppendFormat("+ Carry In: {0:N2}{1}", breakdown("CarryIn"), vbCrLf)
            sb.AppendFormat("- Switch Out: {0:N2}{1}", breakdown("SwitchOut"), vbCrLf)
            sb.AppendFormat("- Balance Out: {0:N2}{1}", breakdown("BalanceOut"), vbCrLf)
            sb.AppendFormat("- Carry Out: {0:N2}{1}", breakdown("CarryOut"), vbCrLf)
            sb.AppendLine("─────────────────────")
            sb.AppendFormat("Total: {0:N2}", breakdown("Total"))

            Return sb.ToString()

        Catch ex As Exception
            Return "Error generating breakdown"
        End Try
    End Function

End Class
