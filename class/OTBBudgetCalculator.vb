Imports System.Data
Imports System.Data.SqlClient

Public Class OTBBudgetCalculator
    Private Shared connectionString As String = ConfigurationManager.ConnectionStrings("BMSConnectionString")?.ConnectionString

    ' --- 1. สร้างตัวแปร Private สำหรับเก็บข้อมูลที่โหลดมา ---
    Private dtOtbTransaction As DataTable
    Private dtSwitchingTransaction As DataTable

    ''' <summary>
    ''' Constructor: เมื่อคลาสนี้ถูก New() จะดึงข้อมูลทั้งหมดจาก DB มาเก็บไว้ก่อน
    ''' </summary>
    Public Sub New()
        LoadAllTransactionData()
    End Sub

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
                           [Brand], [Vendor], [Amount], [Version] 
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
            Dim filter As String = BuildKeyFilter(year, month, category, company, segment, brand, vendor)
            filter &= " AND [Type] = 'Original'"

            Dim originalAmount As Object = dtOtbTransaction.Compute("SUM(Amount)", filter)
            Return If(originalAmount IsNot DBNull.Value, Convert.ToDecimal(originalAmount), 0)

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
            Dim originalAmount As Decimal = GetOriginalAmount(year, month, category, company, segment, brand, vendor)

            Dim filter As String = BuildKeyFilter(year, month, category, company, segment, brand, vendor)
            filter &= " AND [Type] = 'Revise'"

            Dim reviseRows As DataRow() = dtOtbTransaction.Select(filter, "Version ASC")

            Dim totalRevDiff As Decimal = 0
            Dim previousAmount As Decimal = originalAmount

            For Each row As DataRow In reviseRows
                Dim reviseAmount As Decimal = If(row("Amount") IsNot DBNull.Value, Convert.ToDecimal(row("Amount")), 0)
                Dim diff As Decimal = reviseAmount - previousAmount
                totalRevDiff += diff
                previousAmount = reviseAmount
            Next

            Return totalRevDiff
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in GetRevisionDiff (In-Memory): " & ex.Message)
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
            ' ========================================
            ' คำนวณ OUT (จาก Source) - ลบออก
            ' ========================================
            Dim filterOut As String = BuildKeyFilter(year, month, category, company, segment, brand, vendor)
            Dim outRows() As DataRow = dtSwitchingTransaction.Select(filterOut)

            For Each row As DataRow In outRows
                Dim fromType As String = If(row("From") IsNot DBNull.Value, row("From").ToString().Trim().ToUpper(), "")
                Dim amount As Decimal = If(row("BudgetAmount") IsNot DBNull.Value, Convert.ToDecimal(row("BudgetAmount")), 0)

                Select Case fromType
                    Case "D" : result("SwitchOut") += amount
                    Case "G" : result("CarryOut") += amount
                    Case "I" : result("BalanceOut") += amount
                    Case "E" : result("Extra") += amount
                End Select
            Next

            ' ========================================
            ' คำนวณ IN (ไป Destination) - บวกเข้า
            ' ========================================
            Dim filterIn As String = BuildSwitchKeyFilter(year, month, category, company, segment, brand, vendor)
            Dim inRows() As DataRow = dtSwitchingTransaction.Select(filterIn)

            For Each row As DataRow In inRows
                Dim toType As String = If(row("To") IsNot DBNull.Value, row("To").ToString().Trim().ToUpper(), "")
                Dim amount As Decimal = If(row("BudgetAmount") IsNot DBNull.Value, Convert.ToDecimal(row("BudgetAmount")), 0)

                Select Case toType
                    Case "D" : result("SwitchIn") += amount
                    Case "G" : result("CarryIn") += amount
                    Case "I" : result("BalanceIn") += amount
                End Select
            Next

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("Error in CalculateSwitchingAmounts (In-Memory): " & ex.Message)
        End Try

        Return result
    End Function

    ' =================================================================
    ' ===== START: MODIFICATION (แก้ไขจุดที่ Error) ===================
    ' =================================================================

    ''' <summary>
    ''' (ใหม่) Helper Function สำหรับ Escape ' (single quote) สำหรับ .Select Filter
    ''' </summary>
    Private Function EscapeFilter(s As String) As String
        If String.IsNullOrEmpty(s) Then
            Return ""
        End If
        Return s.Replace("'", "''")
    End Function

    ''' <summary>
    ''' (ใหม่) Helper Function สำหรับสร้าง .Select Filter
    ''' </summary>
    Private Function BuildKeyFilter(year As String, month As String, category As String, company As String, segment As String, brand As String, vendor As String) As String
        Return $"[Year] = '{EscapeFilter(year)}' AND " &
               $"[Month] = '{EscapeFilter(month)}' AND " &
               $"[Category] = '{EscapeFilter(category)}' AND " &
               $"[Company] = '{EscapeFilter(company)}' AND " &
               $"[Segment] = '{EscapeFilter(segment)}' AND " &
               $"[Brand] = '{EscapeFilter(brand)}' AND " &
               $"[Vendor] = '{EscapeFilter(vendor)}'"
    End Function

    ''' <summary>
    ''' (ใหม่) Helper Function สำหรับสร้าง .Select Filter (ฝั่ง To/Switch)
    ''' </summary>
    Private Function BuildSwitchKeyFilter(year As String, month As String, category As String, company As String, segment As String, brand As String, vendor As String) As String
        Return $"[SwitchYear] = '{EscapeFilter(year)}' AND " &
               $"[SwitchMonth] = '{EscapeFilter(month)}' AND " &
               $"[SwitchCompany] = '{EscapeFilter(company)}' AND " &
               $"[SwitchCategory] = '{EscapeFilter(category)}' AND " &
               $"[SwitchSegment] = '{EscapeFilter(segment)}' AND " &
               $"[SwitchBrand] = '{EscapeFilter(brand)}' AND " &
               $"[SwitchVendor] = '{EscapeFilter(vendor)}' AND " &
               $"[To] IS NOT NULL"
    End Function

    ' =================================================================
    ' ===== END: MODIFICATION =========================================
    ' =================================================================


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